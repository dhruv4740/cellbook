// Conditional formatting (<conditionalFormatting>/<cfRule>, styles.xml <dxfs>) and data validation
// (<dataValidation>) read/write, as raw OOXML text - dependency-free (no exceljs), shared by both
// the fast reader (fastXlsx.ts) and the exceljs-based reader/writer (xlsx.ts). exceljs's own
// addConditionalFormatting/dataValidations support is missing pieces we need (duplicateValues,
// uniqueValues, timePeriod/dateOccurring formulas, and - critically - dataValidation's
// showDropDown attribute), so both read and write go through here instead for every rule type.
//
// Note the OOXML inversion: a dataValidation's showDropDown="1" HIDES the in-cell dropdown; a
// list validation shows its dropdown by default when the attribute is absent.
import type {
  CellStyle,
  CfDateOccurring,
  CfOperator,
  CfRule,
  CfRuleType,
  DataValidationRule,
  DvAllowType,
  DvErrorStyle,
  DvOperator,
  Range,
  SheetLayout
} from '../model/types'
import { a1, parseA1, rangeA1 } from '../model/types'
import { decodeXmlEntities, extractSection, parseAttrs } from './fastXlsxXml'
import { THEME_COLORS, applyTint } from './xlsxShared'

// ---------- small XML helpers ----------
function escXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function hex6(hex: string | undefined): string {
  const h = (hex || '000000').replace('#', '').toUpperCase()
  return h.length === 6 ? h : h.padStart(6, '0').slice(0, 6)
}
function argb(hex: string | undefined): string {
  return 'FF' + hex6(hex)
}

function resolveColor(attrs: Record<string, string> | undefined): string | undefined {
  if (!attrs) return undefined
  if (attrs.rgb) return '#' + attrs.rgb.slice(-6).toUpperCase()
  if (attrs.theme !== undefined) {
    const idx = parseInt(attrs.theme, 10)
    const base = THEME_COLORS[idx] ?? '000000'
    const tint = attrs.tint !== undefined ? parseFloat(attrs.tint) : 0
    return '#' + (tint ? applyTint(base, tint) : base)
  }
  return undefined
}

/** Finds every non-nested `<tag ...>...</tag>` (or self-closing `<tag/>`) occurrence in `s`. None of
 *  the tags this file looks for (cfRule, formula, dataBar, colorScale, iconSet, cfvo, color, font,
 *  fill, dataValidation, formula1/2, conditionalFormatting, dxf) ever contain another instance of
 *  the same tag, so a flat indexOf scan (like fastXlsx.ts's own element loops) is enough. */
function findAll(s: string, tag: string): { attrs: string; text: string }[] {
  const out: { attrs: string; text: string }[] = []
  let pos = 0
  const openTag = '<' + tag
  const closeTag = '</' + tag + '>'
  const n = s.length
  while (true) {
    const idx = s.indexOf(openTag, pos)
    if (idx === -1 || idx >= n) break
    const after = idx + openTag.length
    const ch = s.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = idx + 1
      continue
    }
    const gt = s.indexOf('>', after)
    if (gt === -1) break
    const selfClose = s.charCodeAt(gt - 1) === 47
    const attrs = s.slice(after, selfClose ? gt - 1 : gt)
    let text = ''
    let next: number
    if (selfClose) {
      next = gt + 1
    } else {
      const close = s.indexOf(closeTag, gt + 1)
      text = close === -1 ? s.slice(gt + 1) : s.slice(gt + 1, close)
      next = close === -1 ? n : close + closeTag.length
    }
    pos = next
    out.push({ attrs, text })
  }
  return out
}

// ============================================================================
// Conditional formatting
// ============================================================================

const DATE_TO_OOXML: Record<CfDateOccurring, string> = {
  yesterday: 'yesterday',
  today: 'today',
  tomorrow: 'tomorrow',
  inTheLast7Days: 'last7Days',
  lastWeek: 'lastWeek',
  thisWeek: 'thisWeek',
  nextWeek: 'nextWeek',
  lastMonth: 'lastMonth',
  thisMonth: 'thisMonth',
  nextMonth: 'nextMonth'
}
const OOXML_TO_DATE: Record<string, CfDateOccurring> = Object.fromEntries(
  Object.entries(DATE_TO_OOXML).map(([k, v]) => [v, k as CfDateOccurring])
)

function topLeft(ranges: Range[]): string {
  const r = ranges[0]
  return r ? a1(r.r1, r.c1) : 'A1'
}

// Formula fallbacks for types Excel evaluates by `type`/`operator`/`timePeriod` alone but still
// ships a formula for (matches what real Excel itself writes, so a file we produce reads the same
// in both this app and real Excel).
function textContainsFormula(rule: CfRule): string {
  const tl = topLeft(rule.ranges)
  const text = (rule.text ?? rule.value ?? '').replace(/"/g, '""')
  return `NOT(ISERROR(SEARCH("${text}",${tl})))`
}
function timePeriodFormula(rule: CfRule): string {
  const tl = topLeft(rule.ranges)
  switch (rule.date) {
    case 'thisWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`
    case 'lastWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`
    case 'nextWeek':
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`
    case 'yesterday':
      return `FLOOR(${tl},1)=TODAY()-1`
    case 'tomorrow':
      return `FLOOR(${tl},1)=TODAY()+1`
    case 'inTheLast7Days':
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`
    case 'lastMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`
    case 'thisMonth':
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`
    case 'nextMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`
    case 'today':
    default:
      return `FLOOR(${tl},1)=TODAY()`
  }
}

const NO_DXF_TYPES: ReadonlySet<CfRuleType> = new Set(['dataBar', 'colorScale', 'iconSet'])

// ---------- Solid Fill data bars (x14 extension) ----------
// A classic (legacy, pre-2010) <dataBar> is always a gradient fill - "Solid Fill" is a newer
// feature that real Excel can only express via the x14 extension list: the legacy <cfRule> keeps
// its plain <dataBar> (so older readers still show *a* bar) plus an <extLst> pointing, by a GUID,
// at a parallel <x14:cfRule> in a single worksheet-level <extLst> that carries gradient="0" and the
// richer x14:dataBar attributes. Both ends have to agree on that GUID and on the sqref.
const X14_MAIN_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main'
const X14_DATABAR_ID_EXT_URI = '{B025F937-C7B1-47D3-B67F-A62EEFF666E3}'
const X14_CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'

function makeGuid(): string {
  const hex = (): string => Math.floor(Math.random() * 16).toString(16).toUpperCase()
  const seg = (n: number): string => Array.from({ length: n }, hex).join('')
  return `{${seg(8)}-${seg(4)}-${seg(4)}-${seg(4)}-${seg(12)}}`
}

function buildCfRuleXml(rule: CfRule, dxfId: number | undefined, solidFillIds: string[]): string {
  const common = `priority="${rule.priority}"${rule.stopIfTrue ? ' stopIfTrue="1"' : ''}`
  const dxfAttr = dxfId !== undefined ? ` dxfId="${dxfId}"` : ''
  switch (rule.type) {
    case 'cellIs': {
      const formulas = [rule.value, rule.value2].filter((v): v is string => v !== undefined)
      return `<cfRule type="cellIs"${dxfAttr} ${common} operator="${rule.operator || 'greaterThan'}">${formulas
        .map((f) => `<formula>${escXml(f)}</formula>`)
        .join('')}</cfRule>`
    }
    case 'textContains': {
      const text = rule.text ?? rule.value ?? ''
      return `<cfRule type="containsText"${dxfAttr} ${common} operator="containsText" text="${escXml(text)}"><formula>${escXml(
        textContainsFormula(rule)
      )}</formula></cfRule>`
    }
    case 'dateOccurring': {
      const tp = DATE_TO_OOXML[rule.date || 'today']
      return `<cfRule type="timePeriod"${dxfAttr} ${common} timePeriod="${tp}"><formula>${escXml(
        timePeriodFormula(rule)
      )}</formula></cfRule>`
    }
    case 'duplicateValues':
      return `<cfRule type="duplicateValues"${dxfAttr} ${common}/>`
    case 'uniqueValues':
      return `<cfRule type="uniqueValues"${dxfAttr} ${common}/>`
    case 'top10':
      return `<cfRule type="top10"${dxfAttr} ${common} percent="${rule.percent ? 1 : 0}" bottom="${
        rule.bottom ? 1 : 0
      }" rank="${rule.rank ?? 10}"/>`
    case 'aboveAverage':
      return `<cfRule type="aboveAverage"${dxfAttr} ${common} aboveAverage="${rule.above === false ? 0 : 1}"${
        rule.equalOrAbove ? ' equalAverage="1"' : ''
      }/>`
    case 'expression': {
      const f = (rule.formula || '').replace(/^=/, '')
      return `<cfRule type="expression"${dxfAttr} ${common}><formula>${escXml(f)}</formula></cfRule>`
    }
    case 'dataBar': {
      const db = rule.dataBar
      const minCfvo = db?.min !== undefined ? `<cfvo type="num" val="${db.min}"/>` : `<cfvo type="min" val="0"/>`
      const maxCfvo = db?.max !== undefined ? `<cfvo type="num" val="${db.max}"/>` : `<cfvo type="max" val="0"/>`
      const showValueAttr = db?.showValue === false ? ' showValue="0"' : ''
      let ext = ''
      if (db?.gradient === false) {
        const guid = makeGuid()
        solidFillIds.push(guid)
        ext = `<extLst><ext xmlns:x14="${X14_MAIN_NS}" uri="${X14_DATABAR_ID_EXT_URI}"><x14:id>${guid}</x14:id></ext></extLst>`
      }
      return `<cfRule type="dataBar" ${common}><dataBar${showValueAttr}>${minCfvo}${maxCfvo}<color rgb="${argb(
        db?.color
      )}"/></dataBar>${ext}</cfRule>`
    }
    case 'colorScale': {
      const cs = rule.colorScale
      const stops = cs?.midColor
        ? `<cfvo type="min" val="0"/><cfvo type="percentile" val="50"/><cfvo type="max" val="0"/><color rgb="${argb(
            cs.minColor
          )}"/><color rgb="${argb(cs.midColor)}"/><color rgb="${argb(cs.maxColor)}"/>`
        : `<cfvo type="min" val="0"/><cfvo type="max" val="0"/><color rgb="${argb(cs?.minColor)}"/><color rgb="${argb(
            cs?.maxColor
          )}"/>`
      return `<cfRule type="colorScale" ${common}><colorScale>${stops}</colorScale></cfRule>`
    }
    case 'iconSet': {
      const is = rule.iconSet
      const name = is?.name || '3TrafficLights1'
      const count = name.startsWith('5') ? 5 : name.startsWith('4') ? 4 : 3
      const cfvos = Array.from({ length: count }, (_, i) => `<cfvo type="percent" val="${Math.round((100 * i) / count)}"/>`).join(
        ''
      )
      const revAttr = is?.reverse ? ' reverse="1"' : ''
      const showAttr = is?.showValue === false ? ' showValue="0"' : ''
      return `<cfRule type="iconSet" ${common}><iconSet iconSet="${escXml(name)}"${revAttr}${showAttr}>${cfvos}</iconSet></cfRule>`
    }
    default:
      return ''
  }
}

function buildX14DataBarExtLst(entries: { guid: string; rule: CfRule; sqref: string }[]): string {
  const inner = entries
    .map(({ guid, rule, sqref }) => {
      const color = argb(rule.dataBar?.color)
      return (
        `<x14:conditionalFormatting xmlns:xm="${XM_NS}">` +
        `<x14:cfRule type="dataBar" id="${guid}">` +
        `<x14:dataBar minLength="0" maxLength="100" gradient="0" border="1" negativeBarBorderColorSameAsPositive="0">` +
        `<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>` +
        `<x14:borderColor rgb="${color}"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/>` +
        `</x14:dataBar></x14:cfRule>` +
        `<xm:sqref>${escXml(sqref)}</xm:sqref>` +
        `</x14:conditionalFormatting>`
      )
    })
    .join('')
  return `<extLst><ext xmlns:x14="${X14_MAIN_NS}" uri="${X14_CF_EXT_URI}"><x14:conditionalFormattings>${inner}</x14:conditionalFormattings></ext></extLst>`
}

/** Builds every `<conditionalFormatting>` block for a sheet, plus (in `x14Ext`) the
 *  worksheet-level x14 extLst any Solid Fill data bars need. `dxfIdFor` assigns (and dedupes)
 *  workbook-global dxf ids for the rule types that reference one. */
export function buildConditionalFormattingXml(
  rules: CfRule[],
  dxfIdFor: (fmt: Partial<CellStyle>) => number
): { xml: string; x14Ext: string } {
  let out = ''
  const solidFill: { guid: string; rule: CfRule; sqref: string }[] = []
  for (const rule of rules) {
    if (!rule.ranges.length) continue
    const sqref = rule.ranges.map(rangeA1).join(' ')
    const dxfId = NO_DXF_TYPES.has(rule.type) ? undefined : dxfIdFor(rule.format || {})
    const solidFillIds: string[] = []
    const ruleXml = buildCfRuleXml(rule, dxfId, solidFillIds)
    if (!ruleXml) continue
    out += `<conditionalFormatting sqref="${escXml(sqref)}">${ruleXml}</conditionalFormatting>`
    for (const guid of solidFillIds) solidFill.push({ guid, rule, sqref })
  }
  return { xml: out, x14Ext: solidFill.length ? buildX14DataBarExtLst(solidFill) : '' }
}

function buildDxfXml(fmt: Partial<CellStyle>): string {
  let font = ''
  if (fmt.color) font += `<color rgb="${argb(fmt.color)}"/>`
  if (fmt.bold) font += '<b/>'
  if (fmt.italic) font += '<i/>'
  if (fmt.underline) font += `<u${fmt.underline === 'double' ? ' val="double"' : ''}/>`
  if (fmt.strike) font += '<strike/>'
  const fill = fmt.fill ? `<fill><patternFill><bgColor rgb="${argb(fmt.fill)}"/></patternFill></fill>` : ''
  const numFmt = fmt.numFmt ? `<numFmt numFmtId="164" formatCode="${escXml(fmt.numFmt)}"/>` : ''
  return `<dxf>${numFmt}${font ? `<font>${font}</font>` : ''}${fill}</dxf>`
}

export function buildDxfsXml(formats: Partial<CellStyle>[]): string {
  if (!formats.length) return ''
  return `<dxfs count="${formats.length}">${formats.map(buildDxfXml).join('')}</dxfs>`
}

function parseDxfEntry(inner: string): Partial<CellStyle> {
  const out: Partial<CellStyle> = {}
  const fontBlock = findAll(inner, 'font')[0]
  if (fontBlock) {
    const f = fontBlock.text
    if (/<b\b[^>]*\/?>/.test(f)) out.bold = true
    if (/<i\b[^>]*\/?>/.test(f)) out.italic = true
    if (/<strike\b[^>]*\/?>/.test(f)) out.strike = true
    const uM = /<u(?:\s+val="([^"]*)")?\s*\/>/.exec(f)
    if (uM) out.underline = uM[1] === 'double' ? 'double' : 'single'
    const colorM = /<color\b([^>]*)\/>/.exec(f)
    if (colorM) {
      const c = resolveColor(parseAttrs(colorM[1], 0, colorM[1].length))
      if (c) out.color = c
    }
  }
  const fillBlock = findAll(inner, 'fill')[0]
  if (fillBlock) {
    const bgM = /<bgColor\b([^>]*)\/>/.exec(fillBlock.text)
    const fgM = /<fgColor\b([^>]*)\/>/.exec(fillBlock.text)
    const colorM = bgM || fgM
    if (colorM) {
      const c = resolveColor(parseAttrs(colorM[1], 0, colorM[1].length))
      if (c) out.fill = c
    }
  }
  const numFmtM = /<numFmt\b([^>]*)\/?>/.exec(inner)
  if (numFmtM) {
    const a = parseAttrs(numFmtM[1], 0, numFmtM[1].length)
    if (a.formatCode) out.numFmt = decodeXmlEntities(a.formatCode)
  }
  return out
}

/** Parses `xl/styles.xml`'s `<dxfs>` list, indexed by dxfId. */
export function parseDxfs(stylesXml: string | undefined): (Partial<CellStyle> | undefined)[] {
  if (!stylesXml) return []
  const section = extractSection(stylesXml, 'dxfs')
  if (!section) return []
  return findAll(section, 'dxf').map((d) => {
    const fmt = parseDxfEntry(d.text)
    return Object.keys(fmt).length ? fmt : undefined
  })
}

/** Reads the worksheet-level x14 extLst's `<x14:cfRule type="dataBar" id="{GUID}">` entries,
 *  mapping each GUID to its `gradient` flag (defaulting to gradient=true, same as a bar with no
 *  x14 extension at all - `gradient="0"` is the marker for Solid Fill). */
function parseX14DataBarGradients(xml: string): Map<string, boolean> {
  const map = new Map<string, boolean>()
  const uriPattern = X14_CF_EXT_URI.replace(/[{}]/g, '\\$&')
  const extM = new RegExp(`<ext\\b[^>]*uri="${uriPattern}"[^>]*>([\\s\\S]*?)</ext>`).exec(xml)
  if (!extM) return map
  for (const cfRule of findAll(extM[1], 'x14:cfRule')) {
    const attrs = parseAttrs(cfRule.attrs, 0, cfRule.attrs.length)
    if (!attrs.id || attrs.type !== 'dataBar') continue
    const dbBlock = findAll(cfRule.text, 'x14:dataBar')[0]
    const dbAttrs = dbBlock ? parseAttrs(dbBlock.attrs, 0, dbBlock.attrs.length) : {}
    map.set(attrs.id, dbAttrs.gradient !== '0')
  }
  return map
}

function parseCfRuleBlock(
  block: { attrs: string; text: string },
  ranges: Range[],
  dxfs: (Partial<CellStyle> | undefined)[],
  fallbackPriority: number,
  x14Gradients: Map<string, boolean>
): CfRule | undefined {
  const attrs = parseAttrs(block.attrs, 0, block.attrs.length)
  const priority = attrs.priority !== undefined ? parseInt(attrs.priority, 10) || fallbackPriority : fallbackPriority
  const stopIfTrue = attrs.stopIfTrue === '1' ? true : undefined
  const dxfId = attrs.dxfId !== undefined ? parseInt(attrs.dxfId, 10) : undefined
  const format = dxfId !== undefined ? dxfs[dxfId] : undefined
  const id = 'cf_' + Math.random().toString(36).slice(2, 9)
  const base = { id, ranges, priority, stopIfTrue, format }
  const formulas = findAll(block.text, 'formula').map((f) => decodeXmlEntities(f.text))

  switch (attrs.type) {
    case 'cellIs':
      return { ...base, type: 'cellIs', operator: attrs.operator as CfOperator, value: formulas[0], value2: formulas[1] }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return { ...base, type: 'textContains', text: attrs.text ?? formulas[0] ?? '' }
    case 'timePeriod':
      return { ...base, type: 'dateOccurring', date: OOXML_TO_DATE[attrs.timePeriod ?? 'today'] ?? 'today' }
    case 'duplicateValues':
      return { ...base, type: 'duplicateValues' }
    case 'uniqueValues':
      return { ...base, type: 'uniqueValues' }
    case 'top10':
      return {
        ...base,
        type: 'top10',
        percent: attrs.percent === '1',
        bottom: attrs.bottom === '1',
        rank: attrs.rank !== undefined ? parseInt(attrs.rank, 10) : 10
      }
    case 'aboveAverage':
      return { ...base, type: 'aboveAverage', above: attrs.aboveAverage !== '0', equalOrAbove: attrs.equalAverage === '1' }
    case 'expression':
      return { ...base, type: 'expression', formula: formulas[0] !== undefined ? '=' + formulas[0] : undefined }
    case 'dataBar': {
      const dbBlock = findAll(block.text, 'dataBar')[0]
      if (!dbBlock) return undefined
      const dbAttrs = parseAttrs(dbBlock.attrs, 0, dbBlock.attrs.length)
      const cfvos = findAll(dbBlock.text, 'cfvo').map((c) => parseAttrs(c.attrs, 0, c.attrs.length))
      const colorEl = findAll(dbBlock.text, 'color')[0]
      const color = (colorEl && resolveColor(parseAttrs(colorEl.attrs, 0, colorEl.attrs.length))) || '#638EC6'
      const min = cfvos[0]?.type === 'num' ? parseFloat(cfvos[0].val) : undefined
      const max = cfvos[1]?.type === 'num' ? parseFloat(cfvos[1].val) : undefined
      let gradient = true
      const extLstBlock = findAll(block.text, 'extLst')[0]
      if (extLstBlock) {
        const idExt = findAll(extLstBlock.text, 'ext').find(
          (e) => parseAttrs(e.attrs, 0, e.attrs.length).uri === X14_DATABAR_ID_EXT_URI
        )
        const guid = idExt && findAll(idExt.text, 'x14:id')[0]?.text.trim()
        if (guid && x14Gradients.has(guid)) gradient = x14Gradients.get(guid)!
      }
      return { ...base, type: 'dataBar', dataBar: { color, min, max, showValue: dbAttrs.showValue !== '0', gradient } }
    }
    case 'colorScale': {
      const csBlock = findAll(block.text, 'colorScale')[0]
      if (!csBlock) return undefined
      const colors = findAll(csBlock.text, 'color').map((c) => resolveColor(parseAttrs(c.attrs, 0, c.attrs.length)) ?? '#FFFFFF')
      const colorScale =
        colors.length >= 3
          ? { minColor: colors[0], midColor: colors[1], maxColor: colors[2] }
          : { minColor: colors[0] ?? '#FFFFFF', maxColor: colors[1] ?? '#FFFFFF' }
      return { ...base, type: 'colorScale', colorScale }
    }
    case 'iconSet': {
      const isBlock = findAll(block.text, 'iconSet')[0]
      if (!isBlock) return undefined
      const isAttrs = parseAttrs(isBlock.attrs, 0, isBlock.attrs.length)
      return {
        ...base,
        type: 'iconSet',
        iconSet: { name: isAttrs.iconSet || '3TrafficLights1', reverse: isAttrs.reverse === '1', showValue: isAttrs.showValue !== '0' }
      }
    }
    default:
      return undefined
  }
}

/** Parses every `<conditionalFormatting sqref="..."><cfRule .../></conditionalFormatting>` block in
 *  a worksheet's raw XML, resolving each rule's dxfId against the workbook's parsed `<dxfs>` list. */
export function parseConditionalFormattings(xml: string, dxfs: (Partial<CellStyle> | undefined)[]): CfRule[] {
  const x14Gradients = parseX14DataBarGradients(xml)
  const out: CfRule[] = []
  let counter = 0
  for (const cf of findAll(xml, 'conditionalFormatting')) {
    const attrs = parseAttrs(cf.attrs, 0, cf.attrs.length)
    if (!attrs.sqref) continue
    const ranges = attrs.sqref
      .split(/\s+/)
      .map((r) => parseA1(r))
      .filter((r): r is Range => r !== null)
    if (!ranges.length) continue
    for (const block of findAll(cf.text, 'cfRule')) {
      counter++
      const rule = parseCfRuleBlock(block, ranges, dxfs, counter, x14Gradients)
      if (rule) out.push(rule)
    }
  }
  return out
}

// ============================================================================
// Data validation
// ============================================================================

const DV_ALLOW_TO_TYPE: Record<DvAllowType, string> = {
  any: 'none',
  whole: 'whole',
  decimal: 'decimal',
  list: 'list',
  date: 'date',
  time: 'time',
  textLength: 'textLength',
  custom: 'custom'
}
const TYPE_TO_DV_ALLOW: Record<string, DvAllowType> = Object.fromEntries(
  Object.entries(DV_ALLOW_TO_TYPE).map(([k, v]) => [v, k as DvAllowType])
)

function dvFormulaOut(rule: DataValidationRule, raw: string | undefined): string | undefined {
  if (raw === undefined) return undefined
  if (rule.allow === 'list') {
    // A range reference is kept with our own leading '=' internally (see getListItems in
    // data/dataValidation.ts); OOXML writes it bare. A literal list is written quoted.
    return raw.startsWith('=') ? raw.slice(1) : `"${raw}"`
  }
  return raw.replace(/^=/, '')
}

function dvFormulaIn(allow: DvAllowType, raw: string | undefined): string | undefined {
  if (raw === undefined) return undefined
  const decoded = decodeXmlEntities(raw)
  if (allow === 'list') {
    return decoded.startsWith('"') && decoded.endsWith('"') ? decoded.slice(1, -1) : '=' + decoded
  }
  // 'custom' is always a formula (the dialog has the person type e.g. "=A1>0"), so keep the leading
  // '=' our own model conventionally carries for it - OOXML itself never writes one.
  if (allow === 'custom') return '=' + decoded
  return decoded
}

const DV_OPERATOR_TYPES: ReadonlySet<DvAllowType> = new Set(['whole', 'decimal', 'date', 'time', 'textLength'])

/** Builds the sheet's `<dataValidations>` element (all rules), or '' when there are none. */
export function buildDataValidationsXml(rules: DataValidationRule[]): string {
  const items = rules
    .map((r) => {
      if (!r.ranges.length) return ''
      const sqref = r.ranges.map(rangeA1).join(' ')
      const type = DV_ALLOW_TO_TYPE[r.allow] ?? 'none'
      let attrs = `type="${type}"`
      if (DV_OPERATOR_TYPES.has(r.allow)) attrs += ` operator="${r.operator || 'between'}"`
      attrs += ` allowBlank="${r.ignoreBlank === false ? 0 : 1}"`
      // showDropDown="1" HIDES the in-cell dropdown - the inversion is intentional (OOXML spec).
      if (r.inCellDropdown === false) attrs += ' showDropDown="1"'
      attrs += ` showInputMessage="${r.showInputMessage ? 1 : 0}"`
      attrs += ` showErrorMessage="${r.showErrorAlert === false ? 0 : 1}"`
      attrs += ` errorStyle="${r.errorStyle || 'stop'}"`
      if (r.errorTitle) attrs += ` errorTitle="${escXml(r.errorTitle)}"`
      if (r.errorMessage) attrs += ` error="${escXml(r.errorMessage)}"`
      if (r.inputTitle) attrs += ` promptTitle="${escXml(r.inputTitle)}"`
      if (r.inputMessage) attrs += ` prompt="${escXml(r.inputMessage)}"`
      attrs += ` sqref="${escXml(sqref)}"`
      const f1 = dvFormulaOut(r, r.formula1)
      const f2 = dvFormulaOut(r, r.formula2)
      const body = (f1 !== undefined ? `<formula1>${escXml(f1)}</formula1>` : '') + (f2 !== undefined ? `<formula2>${escXml(f2)}</formula2>` : '')
      return `<dataValidation ${attrs}>${body}</dataValidation>`
    })
    .filter(Boolean)
    .join('')
  return items ? `<dataValidations count="${rules.length}">${items}</dataValidations>` : ''
}

/** Parses a worksheet's `<dataValidations>` element from its raw XML. */
export function parseDataValidations(xml: string): DataValidationRule[] {
  const section = extractSection(xml, 'dataValidations')
  if (!section) return []
  const out: DataValidationRule[] = []
  let counter = 0
  for (const dv of findAll(section, 'dataValidation')) {
    const attrs = parseAttrs(dv.attrs, 0, dv.attrs.length)
    if (!attrs.sqref) continue
    const ranges = attrs.sqref
      .split(/\s+/)
      .map((r) => parseA1(r))
      .filter((r): r is Range => r !== null)
    if (!ranges.length) continue
    const allow = TYPE_TO_DV_ALLOW[attrs.type ?? 'none'] ?? 'any'
    const f1raw = findAll(dv.text, 'formula1')[0]?.text
    const f2raw = findAll(dv.text, 'formula2')[0]?.text
    out.push({
      id: 'dv_' + counter++ + '_' + Math.random().toString(36).slice(2, 7),
      ranges,
      allow,
      operator: DV_OPERATOR_TYPES.has(allow) ? (attrs.operator as DvOperator | undefined) ?? 'between' : undefined,
      formula1: dvFormulaIn(allow, f1raw),
      formula2: dvFormulaIn(allow, f2raw),
      ignoreBlank: attrs.allowBlank !== '0',
      inCellDropdown: attrs.showDropDown !== '1',
      showInputMessage: attrs.showInputMessage === '1',
      inputTitle: attrs.promptTitle || undefined,
      inputMessage: attrs.prompt || undefined,
      showErrorAlert: attrs.showErrorMessage !== '0',
      errorStyle: (attrs.errorStyle as DvErrorStyle) || 'stop',
      errorTitle: attrs.errorTitle || undefined,
      errorMessage: attrs.error || undefined
    })
  }
  return out
}

// ============================================================================
// Write-side package patching (worksheet XML + styles.xml <dxfs>)
// ============================================================================

// Tags that, per the worksheet schema, come *after* conditionalFormatting/dataValidations - the
// injected block goes right before whichever of these appears first (or before </worksheet>).
const SCHEMA_ANCHORS = [
  '<pageMargins',
  '<pageSetup',
  '<headerFooter',
  '<rowBreaks',
  '<colBreaks',
  '<customProperties',
  '<cellWatches',
  '<ignoredErrors',
  '<smartTags',
  '<drawing',
  '<legacyDrawing',
  '<picture',
  '<oleObjects',
  '<controls',
  '<tableParts',
  '<extLst',
  '</worksheet>'
]

function insertBeforeSchemaAnchor(xml: string, injected: string): string {
  let insertAt = xml.length
  for (const anchor of SCHEMA_ANCHORS) {
    const idx = xml.indexOf(anchor)
    if (idx !== -1 && idx < insertAt) insertAt = idx
  }
  return xml.slice(0, insertAt) + injected + xml.slice(insertAt)
}

/** Ensures the worksheet root element declares `xmlns:x14` and lists `x14` in `mc:Ignorable`,
 *  matching what real Excel adds whenever a file carries an x14 extLst (needed for the Solid Fill
 *  data bar extension below). */
function ensureX14Namespace(xml: string): string {
  const openIdx = xml.indexOf('<worksheet')
  if (openIdx === -1) return xml
  const gt = xml.indexOf('>', openIdx)
  if (gt === -1) return xml
  let tag = xml.slice(openIdx, gt)
  if (!/xmlns:x14=/.test(tag)) tag += ` xmlns:x14="${X14_MAIN_NS}"`
  if (/mc:Ignorable="/.test(tag)) {
    tag = tag.replace(/mc:Ignorable="([^"]*)"/, (_m, val: string) => {
      const tokens = val.split(/\s+/).filter(Boolean)
      if (!tokens.includes('x14')) tokens.push('x14')
      return `mc:Ignorable="${tokens.join(' ')}"`
    })
  } else {
    tag += ' mc:Ignorable="x14"'
    if (!/xmlns:mc=/.test(tag)) tag += ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  }
  return xml.slice(0, openIdx) + tag + xml.slice(gt)
}

/** Appends a worksheet-level `<extLst>...</extLst>` block (must be the worksheet's very last
 *  child per the schema, after even pageMargins/drawing/tableParts) - merging into an existing
 *  *worksheet-level* extLst if the sheet already has one, else inserting a fresh one right before
 *  `</worksheet>`. Anchored on "immediately before </worksheet>" rather than "the first <extLst>
 *  anywhere" - a Solid Fill data bar's own per-cfRule <extLst> (nested inside
 *  <conditionalFormatting>, earlier in the document) is a different element and must not be matched
 *  here. */
function appendWorksheetExtLst(xml: string, extBlock: string): string {
  const existingM = /<extLst>([\s\S]*?)<\/extLst>\s*<\/worksheet>/.exec(xml)
  const innerToAdd = extBlock.slice('<extLst>'.length, extBlock.length - '</extLst>'.length)
  if (existingM) {
    const extLstStart = existingM.index
    const extLstFullLen = '<extLst>'.length + existingM[1].length + '</extLst>'.length
    return xml.slice(0, extLstStart) + `<extLst>${existingM[1]}${innerToAdd}</extLst>` + xml.slice(extLstStart + extLstFullLen)
  }
  const closeIdx = xml.lastIndexOf('</worksheet>')
  if (closeIdx === -1) return xml
  return xml.slice(0, closeIdx) + extBlock + xml.slice(closeIdx)
}

function insertDxfsIntoStyles(stylesXml: string, dxfsXml: string): string {
  if (/<dxfs\b/.test(stylesXml)) {
    // Merge into an existing (e.g. exceljs-authored, always-empty-in-practice) <dxfs> block.
    return stylesXml.replace(/<dxfs\b[^>]*\/>/, dxfsXml).replace(/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/, dxfsXml)
  }
  const idx = stylesXml.indexOf('<tableStyles')
  if (idx !== -1) return stylesXml.slice(0, idx) + dxfsXml + stylesXml.slice(idx)
  const closeIdx = stylesXml.indexOf('</styleSheet>')
  if (closeIdx === -1) return stylesXml
  return stylesXml.slice(0, closeIdx) + dxfsXml + stylesXml.slice(closeIdx)
}

/** Patches conditional formatting + data validation into an already-unzipped package's files
 *  (worksheet XML and, when any rule needs one, `xl/styles.xml`'s `<dxfs>`). Mutates `files`. */
export function patchCfAndDv(files: Record<string, Uint8Array>, sheetLayouts: (Partial<SheetLayout> | undefined)[]): void {
  const dec = new TextDecoder('utf-8')
  const enc = new TextEncoder()

  const dxfList: Partial<CellStyle>[] = []
  const dxfKeyToId = new Map<string, number>()
  const dxfIdFor = (fmt: Partial<CellStyle>): number => {
    const key = JSON.stringify(fmt)
    let id = dxfKeyToId.get(key)
    if (id === undefined) {
      id = dxfList.length
      dxfList.push(fmt)
      dxfKeyToId.set(key, id)
    }
    return id
  }

  sheetLayouts.forEach((layout, i) => {
    const cf = layout?.conditionalFormats ?? []
    const dv = layout?.dataValidations ?? []
    if (!cf.length && !dv.length) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    const bytes = files[path]
    if (!bytes) return
    let xml = dec.decode(bytes)
    const { xml: cfXml, x14Ext } = buildConditionalFormattingXml(cf, dxfIdFor)
    const dvXml = buildDataValidationsXml(dv)
    const injected = cfXml + dvXml
    let changed = false
    if (injected) {
      xml = insertBeforeSchemaAnchor(xml, injected)
      changed = true
    }
    if (x14Ext) {
      xml = appendWorksheetExtLst(ensureX14Namespace(xml), x14Ext)
      changed = true
    }
    if (changed) files[path] = enc.encode(xml)
  })

  if (dxfList.length) {
    const stylesPath = 'xl/styles.xml'
    const stylesBytes = files[stylesPath]
    if (stylesBytes) {
      const patched = insertDxfsIntoStyles(dec.decode(stylesBytes), buildDxfsXml(dxfList))
      files[stylesPath] = enc.encode(patched)
    }
  }
}

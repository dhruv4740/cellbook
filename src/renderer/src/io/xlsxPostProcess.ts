import { colName } from '../model/types'
// Post-processing on the raw zip package that exceljs's object model can't express: per-column
// AutoFilter criteria (exceljs only round-trips the filtered range, never filterColumn/filters)
// and, for .xlsm saves, the VBA project part exceljs always strips. Runs on the bytes
// wb.xlsx.writeBuffer() already produced, so a normal .xlsx save with no criteria and no macros
// pays nothing beyond the unzip/rezip when there's actually something to patch.
import { unzipSync, zipSync } from 'fflate'
import type { AutoFilter, CalcMode, ChartSpec, DataValidationRule, SheetLayout, SheetProtection, WorkbookProtection } from '../model/types'
import { passwordAttrs, sheetProtectionXml } from '../model/protection'
import { outlineRuns } from '../model/outline'
import { buildPageSetupXml } from './pageSetupXml'
import type { PageSetup } from '../model/types'
import { rangeA1 } from '../model/types'
import { BLANK_TOKEN } from '../engine/filter'
import { parseRelationships, resolveRelTarget } from './fastXlsxXml'
import { buildChartSpaceXml, buildTwoCellAnchorXml } from './xlsxCharts'

export interface VbaProject {
  bin: Uint8Array
  signature?: Uint8Array
  /** Raw bytes of the source workbook's xl/_rels/vbaProject.bin.rels, copied through verbatim -
   *  simplest way to keep the signature part's own relationship intact. */
  signatureRels?: Uint8Array
}

const VBA_REL_TYPE = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeXmlText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/** Builds the full `<autoFilter ref="..">...</autoFilter>` element, replacing exceljs's
 *  range-only self-closing tag. colId is relative to the range's first column per the OOXML spec. */
function buildAutoFilterXml(af: AutoFilter): string {
  const ref = rangeA1(af.range)
  const body = Object.entries(af.columns)
    .map(([colStr, crit]) => {
      const colId = Number(colStr) - af.range.c1
      if (colId < 0) return ''
      const hasBlank = crit.selected.includes(BLANK_TOKEN)
      const filters = crit.selected
        .filter((v) => v !== BLANK_TOKEN)
        .map((v) => `<filter val="${escapeXmlAttr(v)}"/>`)
        .join('')
      return `<filterColumn colId="${colId}"><filters${hasBlank ? ' blank="1"' : ''}>${filters}</filters></filterColumn>`
    })
    .join('')
  return `<autoFilter ref="${ref}">${body}</autoFilter>`
}

function buildDataValidationXml(rule: DataValidationRule): string {
  const attrs: string[] = []

  if (rule.allow && rule.allow !== 'any') {
    attrs.push(`type="${escapeXmlAttr(rule.allow)}"`)
  }

  if (rule.errorStyle) {
    attrs.push(`errorStyle="${escapeXmlAttr(rule.errorStyle)}"`)
  }

  if (
    rule.operator &&
    rule.allow !== 'list' &&
    rule.allow !== 'custom' &&
    rule.allow !== 'any'
  ) {
    attrs.push(`operator="${escapeXmlAttr(rule.operator)}"`)
  }

  if (rule.ignoreBlank !== undefined) {
    attrs.push(`allowBlank="${rule.ignoreBlank ? '1' : '0'}"`)
  }

  // Note Excel's inverted flag: showDropDown="1" means the in-cell dropdown is HIDDEN.
  if (rule.allow === 'list' && rule.inCellDropdown === false) {
    attrs.push('showDropDown="1"')
  }

  if (rule.showInputMessage !== undefined) {
    attrs.push(`showInputMessage="${rule.showInputMessage ? '1' : '0'}"`)
  }

  if (rule.showErrorAlert !== undefined) {
    attrs.push(`showErrorMessage="${rule.showErrorAlert ? '1' : '0'}"`)
  }

  if (rule.inputTitle) {
    attrs.push(`promptTitle="${escapeXmlAttr(rule.inputTitle)}"`)
  }

  if (rule.inputMessage) {
    attrs.push(`prompt="${escapeXmlAttr(rule.inputMessage)}"`)
  }

  if (rule.errorTitle) {
    attrs.push(`errorTitle="${escapeXmlAttr(rule.errorTitle)}"`)
  }

  if (rule.errorMessage) {
    attrs.push(`error="${escapeXmlAttr(rule.errorMessage)}"`)
  }

  const sqref = rule.ranges.map(rangeA1).join(' ')
  attrs.push(`sqref="${escapeXmlAttr(sqref)}"`)

  let children = ''
  if (rule.formula1 !== undefined && rule.formula1 !== '') {
    let f1 = rule.formula1
    if (rule.allow === 'list' && !f1.startsWith('=') && !f1.startsWith('"')) {
      f1 = `"${f1}"`
    }
    children += `<formula1>${escapeXmlText(f1)}</formula1>`
  }
  if (rule.formula2 !== undefined && rule.formula2 !== '') {
    children += `<formula2>${escapeXmlText(rule.formula2)}</formula2>`
  }

  const attrStr = attrs.length ? ' ' + attrs.join(' ') : ''
  if (children) {
    return `<dataValidation${attrStr}>${children}</dataValidation>`
  }
  return `<dataValidation${attrStr}/>`
}

export function buildDataValidationsXml(rules: DataValidationRule[]): string {
  if (!rules || !rules.length) return ''
  const items = rules.map(buildDataValidationXml).join('')
  return `<dataValidations count="${rules.length}">${items}</dataValidations>`

}

/** ExcelJS puts <legacyDrawing> (notes) after a data bar's worksheet <extLst>; Excel requires
 *  extLst last and refuses to open the file otherwise. Moves it back; returns the input if nothing moved. */
export function fixLegacyDrawingOrder(zipBytes: Uint8Array): Uint8Array {
  const files = unzipSync(zipBytes)
  let changed = false
  for (const name of Object.keys(files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue
    const xml = new TextDecoder().decode(files[name])
    const ext = execTopLevel(/<extLst>/g, xml)
    const legacy = /<legacyDrawing\b[^>]*\/>/.exec(xml)
    if (!ext || !legacy || legacy.index < ext.index) continue
    const without = xml.slice(0, legacy.index) + xml.slice(legacy.index + legacy[0].length)
    files[name] = new TextEncoder().encode(without.slice(0, ext.index) + legacy[0] + without.slice(ext.index))
    changed = true
  }
  return changed ? zipSync(files, { level: 6 }) : zipBytes
}

const DYNAMIC_ARRAY_METADATA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'

/** Marks array formulas as dynamic (spilling) arrays the way Excel does: cm="1" on the cell plus
 *  xl/metadata.xml. Without it Excel treats FILTER/UNIQUE as fixed legacy array formulas. */
export function markDynamicArrays(zipBytes: Uint8Array): Uint8Array {
  const files = unzipSync(zipBytes)
  const dec = new TextDecoder()
  const enc = new TextEncoder()
  let any = false
  for (const name of Object.keys(files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue
    const xml = dec.decode(files[name])
    if (xml.indexOf('t="array"') === -1) continue
    const next = xml.replace(/<c ((?:(?!cm=)[^>])*?)>(<f t="array")/g, '<c $1 cm="1">$2')
    if (next !== xml) { files[name] = enc.encode(next); any = true }
  }
  if (!any) return zipBytes
  if (!files['xl/metadata.xml']) {
    files['xl/metadata.xml'] = enc.encode(DYNAMIC_ARRAY_METADATA)
    const ct = dec.decode(files['[Content_Types].xml'])
    files['[Content_Types].xml'] = enc.encode(ct.replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>'))
    const rels = dec.decode(files['xl/_rels/workbook.xml.rels'])
    files['xl/_rels/workbook.xml.rels'] = enc.encode(rels.replace('</Relationships>', '<Relationship Id="rIdMeta1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>'))
  }
  return zipSync(files, { level: 6 })
}

/** Finds a worksheet-level element (global regex) past sheetData/conditionalFormatting, so an
 *  <extLst> nested inside a data-bar cfRule is never mistaken for the sheet's own extLst. */
function execTopLevel(re: RegExp, xml: string): RegExpExecArray | null {
  re.lastIndex = Math.max(0, xml.lastIndexOf('</sheetData>'), xml.lastIndexOf('</conditionalFormatting>'))
  return re.exec(xml)
}

function addContentTypeOverride(ctXml: string, partName: string, contentType: string): string {
  if (ctXml.includes(`PartName="${partName}"`)) return ctXml
  const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`
  return ctXml.replace('</Types>', override + '</Types>')
}

function injectCharts(
  files: Record<string, Uint8Array>,
  sheetCharts: (ChartSpec[] | undefined)[],
  sheetNames: string[],
  sheetCells: Record<string, { raw?: string }>[],
  sheetLayouts?: (Partial<SheetLayout> | undefined)[]
): void {
  const dec = new TextDecoder('utf-8')
  const enc = new TextEncoder()

  const ctPath = '[Content_Types].xml'
  let ctXml = dec.decode(files[ctPath])
  if (!ctXml.includes('Extension="xml"')) {
    ctXml = ctXml.replace('<Types', '<Types><Default Extension="xml" ContentType="application/xml"/>')
  }

  let maxDrawingNum = 0
  let maxChartNum = 0
  for (const key of Object.keys(files)) {
    const mD = /^xl\/drawings\/drawing(\d+)\.xml$/.exec(key)
    if (mD) maxDrawingNum = Math.max(maxDrawingNum, parseInt(mD[1], 10))
    const mC = /^xl\/charts\/chart(\d+)\.xml$/.exec(key)
    if (mC) maxChartNum = Math.max(maxChartNum, parseInt(mC[1], 10))
  }

  for (let i = 0; i < sheetCharts.length; i++) {
    const charts = sheetCharts[i]
    if (!charts || !charts.length) continue

    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
    const sheetBytes = files[sheetPath]
    if (!sheetBytes) continue
    let sheetXml = dec.decode(sheetBytes)
    const sheetName = sheetNames[i] || `Sheet${i + 1}`
    const cells = sheetCells[i] || {}
    const cellReader = {
      get: (r: number, c: number) => {
        const raw = cells[r + ',' + c]?.raw
        if (raw === undefined || raw === '') return null
        if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw)) return Number(raw)
        if (raw === 'TRUE') return true
        if (raw === 'FALSE') return false
        return raw
      }
    }

    const sheetRelsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let sheetRelsXml = files[sheetRelsPath] ? dec.decode(files[sheetRelsPath]) : ''
    let drawingRId = ''
    let drawingPath = ''

    const drawMatch = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)
    if (drawMatch) {
      drawingRId = drawMatch[1]
      const rels = parseRelationships(sheetRelsXml)
      const r = rels.find((rel) => rel.id === drawingRId)
      if (r) drawingPath = resolveRelTarget(sheetPath, r.target)
    }

    if (!drawingPath) {
      maxDrawingNum++
      const drawingFile = `drawing${maxDrawingNum}.xml`
      drawingPath = `xl/drawings/${drawingFile}`

      let maxRId = 0
      for (const m of sheetRelsXml.matchAll(/Id="rId(\d+)"/g)) {
        maxRId = Math.max(maxRId, parseInt(m[1], 10))
      }
      drawingRId = `rId${maxRId + 1}`

      const newRel = `<Relationship Id="${drawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFile}"/>`
      if (sheetRelsXml.includes('</Relationships>')) {
        sheetRelsXml = sheetRelsXml.replace('</Relationships>', newRel + '</Relationships>')
      } else {
        sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${newRel}\n</Relationships>`
      }
      files[sheetRelsPath] = enc.encode(sheetRelsXml)

      const drawingXmlTag = `<drawing r:id="${drawingRId}"/>`
      const insertBeforeRegex = /<(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s\/>]|<\/worksheet>/g
      const m = execTopLevel(insertBeforeRegex, sheetXml)
      if (m) {
        sheetXml = sheetXml.slice(0, m.index) + drawingXmlTag + sheetXml.slice(m.index)
      } else {
        sheetXml = sheetXml.replace('</worksheet>', drawingXmlTag + '</worksheet>')
      }
      files[sheetPath] = enc.encode(sheetXml)
    }

    const drawingDir = drawingPath.split('/').slice(0, -1).join('/')
    const drawingFileName = drawingPath.split('/').pop() ?? ''
    const drawingRelsPath = `${drawingDir}/_rels/${drawingFileName}.rels`
    let drawingRelsXml = files[drawingRelsPath] ? dec.decode(files[drawingRelsPath]) : ''
    let drawingXml = files[drawingPath] ? dec.decode(files[drawingPath]) : ''

    let maxDrawingRId = 0
    for (const m of drawingRelsXml.matchAll(/Id="rId(\d+)"/g)) {
      maxDrawingRId = Math.max(maxDrawingRId, parseInt(m[1], 10))
    }

    const anchors: string[] = []
    const drawingRelsToAdd: string[] = []

    for (let cIdx = 0; cIdx < charts.length; cIdx++) {
      const spec = charts[cIdx]
      maxChartNum++
      const chartNum = maxChartNum
      const chartPath = `xl/charts/chart${chartNum}.xml`

      maxDrawingRId++
      const chartRelId = `rId${maxDrawingRId}`

      const chartXml = buildChartSpaceXml(spec, sheetName, cellReader)
      files[chartPath] = enc.encode(chartXml)

      const anchorXml = buildTwoCellAnchorXml(spec, chartRelId, cIdx, sheetLayouts?.[i])
      anchors.push(anchorXml)

      drawingRelsToAdd.push(
        `<Relationship Id="${chartRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNum}.xml"/>`
      )

      ctXml = addContentTypeOverride(
        ctXml,
        `/${chartPath}`,
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
      )
    }

    ctXml = addContentTypeOverride(
      ctXml,
      `/${drawingPath}`,
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    )

    if (drawingRelsXml.includes('</Relationships>')) {
      drawingRelsXml = drawingRelsXml.replace(
        '</Relationships>',
        drawingRelsToAdd.join('') + '</Relationships>'
      )
    } else {
      drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${drawingRelsToAdd.join('')}\n</Relationships>`
    }
    files[drawingRelsPath] = enc.encode(drawingRelsXml)

    if (drawingXml.includes('</xdr:wsDr>')) {
      drawingXml = drawingXml.replace('</xdr:wsDr>', anchors.join('') + '</xdr:wsDr>')
    } else {
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">\n${anchors.join('')}\n</xdr:wsDr>`
    }
    files[drawingPath] = enc.encode(drawingXml)
  }

  files[ctPath] = enc.encode(ctXml)
}

function injectVbaProject(files: Record<string, Uint8Array>, vba: VbaProject): void {
  files['xl/vbaProject.bin'] = vba.bin
  if (vba.signature) files['xl/vbaProjectSignature.bin'] = vba.signature
  if (vba.signatureRels) files['xl/_rels/vbaProject.bin.rels'] = vba.signatureRels

  const dec = new TextDecoder('utf-8')
  const enc = new TextEncoder()

  // [Content_Types].xml: a .bin default (covers vbaProject.bin and vbaProjectSignature.bin) plus
  // the macro-enabled content type for the workbook part itself.
  const ctPath = '[Content_Types].xml'
  let ct = dec.decode(files[ctPath])
  if (!/Extension="bin"/.test(ct)) {
    ct = ct.replace(
      '<Default Extension="rels"',
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="rels"'
    )
  }
  ct = ct.replace(
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
    'ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"'
  )
  files[ctPath] = enc.encode(ct)

  // xl/_rels/workbook.xml.rels: point the workbook part at the VBA project part.
  const relsPath = 'xl/_rels/workbook.xml.rels'
  const rels = dec.decode(files[relsPath])
  let maxId = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, parseInt(m[1], 10))
  const rel = `<Relationship Id="rId${maxId + 1}" Type="${VBA_REL_TYPE}" Target="vbaProject.bin"/>`
  files[relsPath] = enc.encode(rels.replace('</Relationships>', rel + '</Relationships>'))
}

/** `sheetFilters[i]` is sheet `i+1`'s AutoFilter (exceljs names worksheet parts sheet1.xml,
 *  sheet2.xml, ... in the order worksheets were added, which matches WorkbookData.sheets order). */
export function postProcessXlsx(
  zipBytes: Uint8Array,
  sheetFilters: (AutoFilter | undefined)[],
  vba?: VbaProject,
  sheetValidations?: (DataValidationRule[] | undefined)[],
  sheetCharts?: (ChartSpec[] | undefined)[],
  sheetNames?: string[],
  sheetCells?: Record<string, { raw?: string }>[],
  sheetLayouts?: (Partial<SheetLayout> | undefined)[]
): Uint8Array {
  const files = unzipSync(zipBytes)
  const sheetCount = Math.max(sheetFilters.length, sheetValidations?.length ?? 0)

  for (let i = 0; i < sheetCount; i++) {
    const af = sheetFilters[i]
    const dvs = sheetValidations?.[i]
    const hasAf = af && Object.keys(af.columns).length > 0
    const hasDvs = dvs && dvs.length > 0
    if (!hasAf && !hasDvs) continue

    const path = `xl/worksheets/sheet${i + 1}.xml`
    const xml = files[path]
    if (!xml) continue
    let text = new TextDecoder('utf-8').decode(xml)

    if (hasAf) {
      text = text.replace(/<autoFilter ref="[^"]*"\/>/, buildAutoFilterXml(af!))
    }

    if (hasDvs) {
      const dvXml = buildDataValidationsXml(dvs!)
      if (dvXml) {
        if (text.includes('<dataValidations')) {
          text = text.replace(/<dataValidations[\s\S]*?<\/dataValidations>/, dvXml)
        } else {
          const insertBeforeRegex = /<(?:hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s\/>]|<\/worksheet>/g
          const m = execTopLevel(insertBeforeRegex, text)
          if (m) {
            text = text.slice(0, m.index) + dvXml + text.slice(m.index)
          } else {
            text = text.replace('</worksheet>', dvXml + '</worksheet>')
          }
        }
      }
    }

    files[path] = new TextEncoder().encode(text)
  }

  if (sheetCharts && sheetCharts.some((c) => c && c.length > 0)) {
    injectCharts(files, sheetCharts, sheetNames || [], sheetCells || [], sheetLayouts)
  }

  if (vba) injectVbaProject(files, vba)

  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) {
    if (name.endsWith('/')) continue // directory placeholder, not a real entry we need to keep
    out[name] = bytes
  }
  return zipSync(out, { level: 6 })
}

/** Writes the workbook's defined names into xl/workbook.xml, keeping any exceljs already wrote
 *  (e.g. _xlnm._FilterDatabase) unless ours has the same name and scope. */
export function injectDefinedNames(zipBytes: Uint8Array, names: { name: string; ref: string; sheet?: number; hidden?: boolean }[]): Uint8Array {
  if (!names.length) return zipBytes
  const files = unzipSync(zipBytes)
  const path = 'xl/workbook.xml'
  let xml = new TextDecoder('utf-8').decode(files[path])
  const key = (n: string, s?: number) => n.toLowerCase() + '|' + (s ?? '')
  const ours = new Set(names.map((n) => key(n.name, n.sheet)))
  const kept: string[] = []
  const existing = /<definedNames>([\s\S]*?)<\/definedNames>/.exec(xml)
  if (existing) {
    for (const m of existing[1].matchAll(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g)) {
      const nm = /name="([^"]*)"/.exec(m[0])?.[1] ?? ''
      const ls = /localSheetId="(\d+)"/.exec(m[0])?.[1]
      if (!ours.has(key(nm, ls !== undefined ? +ls : undefined))) kept.push(m[0])
    }
  }
  const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const mine = names.map((n) => `<definedName name="${esc(n.name)}"${n.sheet !== undefined ? ` localSheetId="${n.sheet}"` : ''}${n.hidden ? ' hidden="1"' : ''}>${esc(n.ref)}</definedName>`)
  const block = `<definedNames>${[...kept, ...mine].join('')}</definedNames>`
  xml = existing ? xml.replace(existing[0], block) : xml.replace('</sheets>', '</sheets>' + block)
  files[path] = new TextEncoder().encode(xml)
  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) if (!name.endsWith('/')) out[name] = bytes
  return zipSync(out, { level: 6 })
}

/** Writes each sheet's hyperlinks (sheet XML <hyperlinks> + external relationships). */
export function injectHyperlinks(zipBytes: Uint8Array, sheetLinks: (Record<string, { target?: string; location?: string; tooltip?: string }> | undefined)[]): Uint8Array {
  if (!sheetLinks.some((l) => l && Object.keys(l).length)) return zipBytes
  const files = unzipSync(zipBytes)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  sheetLinks.forEach((links, i) => {
    if (!links || !Object.keys(links).length) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    if (!files[path]) return
    let xml = dec.decode(files[path])
    if (xml.includes('<hyperlinks')) return // already written (e.g. by exceljs): keep it
    const relsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let rels = files[relsPath] ? dec.decode(files[relsPath]) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    let maxId = 0
    for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, +m[1])
    const items: string[] = []
    const newRels: string[] = []
    for (const [key, link] of Object.entries(links)) {
      const [r, c] = key.split(',').map(Number)
      const ref = colName(c) + (r + 1)
      let attrs = `ref="${ref}"`
      if (link.target) {
        const id = `rId${++maxId}`
        newRels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(link.target)}" TargetMode="External"/>`)
        attrs += ` r:id="${id}"`
      }
      if (link.location) attrs += ` location="${esc(link.location)}"`
      if (link.tooltip) attrs += ` tooltip="${esc(link.tooltip)}"`
      items.push(`<hyperlink ${attrs}/>`)
    }
    const block = `<hyperlinks>${items.join('')}</hyperlinks>`
    // Schema order: hyperlinks come after dataValidations and before printOptions/pageMargins.
    const m = execTopLevel(/<(?:printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s/>]|<\/worksheet>/g, xml)
    xml = m ? xml.slice(0, m.index) + block + xml.slice(m.index) : xml
    if (!/xmlns:r=/.test(xml.slice(0, 600))) xml = xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
    files[path] = enc.encode(xml)
    if (newRels.length) {
      rels = rels.replace('</Relationships>', newRels.join('') + '</Relationships>')
      files[relsPath] = enc.encode(rels)
    }
  })
  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) if (!name.endsWith('/')) out[name] = bytes
  return zipSync(out, { level: 6 })
}

/** Writes sheet protection, workbook protection and the calculation mode - none of which exceljs can express faithfully
 *  (it re-hashes passwords and inverts the option flags). Placement follows the schema: sheetProtection right after sheetData. */
export function injectProtection(
  zipBytes: Uint8Array,
  sheetProt: (SheetProtection | undefined)[],
  wbProt: WorkbookProtection | undefined,
  calcMode: CalcMode | undefined
): Uint8Array {
  if (!sheetProt.some(Boolean) && !wbProt && (!calcMode || calcMode === 'auto')) return zipBytes
  const files = unzipSync(zipBytes)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  sheetProt.forEach((p, i) => {
    if (!p) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    if (!files[path]) return
    let xml = dec.decode(files[path])
    xml = xml.replace(/<sheetProtection\b[^>]*\/>/, '')
    const frag = sheetProtectionXml(p)
    if (xml.includes('</sheetData>')) xml = xml.replace('</sheetData>', '</sheetData>' + frag)
    else xml = xml.replace('<sheetData/>', '<sheetData/>' + frag)
    files[path] = enc.encode(xml)
  })
  let wx = dec.decode(files['xl/workbook.xml'])
  if (wbProt && (wbProt.lockStructure || wbProt.lockWindows)) {
    const frag = `<workbookProtection${passwordAttrs(wbProt.password, 'workbook')}${wbProt.lockStructure ? ' lockStructure="1"' : ''}${wbProt.lockWindows ? ' lockWindows="1"' : ''}/>`
    wx = wx.replace(/<workbookProtection\b[^>]*\/>/, '')
    wx = wx.includes('<bookViews') ? wx.replace('<bookViews', frag + '<bookViews') : wx.replace('<sheets>', frag + '<sheets>')
  }
  if (calcMode && calcMode !== 'auto') {
    const calc = calcMode === 'manual' ? 'calcMode="manual"' : 'calcMode="autoNoTable"'
    if (/<calcPr\b/.test(wx)) wx = wx.replace(/<calcPr\b([^>]*?)\/>/, (_m, a: string) => `<calcPr${a.replace(/\s*calcMode="[^"]*"/, '')} ${calc}/>`)
    else wx = wx.replace('</workbook>', `<calcPr ${calc}/></workbook>`)
  }
  files['xl/workbook.xml'] = enc.encode(wx)
  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) if (!name.endsWith('/')) out[name] = bytes
  return zipSync(out, { level: 6 })
}

/** Writes Excel's `collapsed` flags on the summary rows/columns of collapsed groups (and nowhere else). */
export function injectOutline(zipBytes: Uint8Array, layouts: (Partial<SheetLayout> | undefined)[]): Uint8Array {
  if (!layouts.some((l) => l && (l.rowLevels || l.colLevels))) return zipBytes
  const files = unzipSync(zipBytes)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  layouts.forEach((l, i) => {
    if (!l || (!l.rowLevels && !l.colLevels)) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    if (!files[path]) return
    let xml = dec.decode(files[path])
    xml = xml.replace(/(<(?:row|col)\b[^>]*?) collapsed="1"/g, '$1')
    const full = { colWidths: {}, rowHeights: {}, hiddenRows: {}, hiddenCols: {}, merges: [], frozen: { rows: 0, cols: 0 }, showGridlines: true, showHeadings: true, zoom: 100, notes: {}, comments: {}, ...l } as SheetLayout
    for (const run of outlineRuns(full, 'row')) {
      if (!run.collapsed || run.summary < 0) continue
      const n = run.summary + 1
      xml = xml.replace(new RegExp(`<row r="${n}"(?=[ >/])`), `<row r="${n}" collapsed="1"`)
    }
    for (const run of outlineRuns(full, 'col')) {
      if (!run.collapsed || run.summary < 0) continue
      const n = run.summary + 1
      xml = xml.replace(new RegExp(`<col min="${n}" max="${n}"(?=[ >/])`), `<col min="${n}" max="${n}" collapsed="1"`)
    }
    files[path] = enc.encode(xml)
  })
  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) if (!name.endsWith('/')) out[name] = bytes
  return zipSync(out, { level: 6 })
}

/** Replaces exceljs's placeholder page elements with the sheet's real Page Setup (margins, header/footer, scale, breaks...). */
export function injectPageSetup(zipBytes: Uint8Array, setups: (PageSetup | undefined)[]): Uint8Array {
  if (!setups.some(Boolean)) return zipBytes
  const files = unzipSync(zipBytes)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  setups.forEach((p, i) => {
    if (!p) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    if (!files[path]) return
    let xml = dec.decode(files[path])
    xml = xml.replace(/<printOptions\b[^>]*\/>/, '').replace(/<pageMargins\b[^>]*\/>/, '').replace(/<pageSetup\b[^>]*\/>/, '')
    xml = xml.replace(/<headerFooter\b[^>]*(?:\/>|>[\s\S]*?<\/headerFooter>)/, '').replace(/<rowBreaks\b[\s\S]*?<\/rowBreaks>/, '').replace(/<colBreaks\b[\s\S]*?<\/colBreaks>/, '')
    // fit-to-page lives in sheetPr/pageSetUpPr (the last child of sheetPr)
    xml = xml.replace(/<pageSetUpPr\b[^>]*\/>/, '')
    if (p.fitToPage) {
      const pr = '<pageSetUpPr fitToPage="1"/>'
      if (/<sheetPr\b[^>]*>/.test(xml) && !/<sheetPr\b[^>]*\/>/.test(xml)) xml = xml.replace('</sheetPr>', pr + '</sheetPr>')
      else xml = xml.replace(/<sheetPr\b[^>]*\/>/, '').replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr>' + pr + '</sheetPr>')
    }
    const frag = buildPageSetupXml(p)
    // Schema order: these come after hyperlinks/dataValidations and before drawing, legacyDrawing, tableParts, extLst.
    const m = execTopLevel(/<(?:customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[ />]|<\/worksheet>/g, xml)
    xml = m ? xml.slice(0, m.index) + frag + xml.slice(m.index) : xml
    files[path] = enc.encode(xml)
  })
  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) if (!name.endsWith('/')) out[name] = bytes
  return zipSync(out, { level: 6 })
}

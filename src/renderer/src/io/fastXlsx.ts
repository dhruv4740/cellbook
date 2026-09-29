// Fast .xlsx reader: fflate `unzipSync` + a hand-written XML scanner (see fastXlsxXml.ts), no
// exceljs and no DOMParser. Produces exactly the WorkbookData shape/px conversions that
// readXlsxExceljs (xlsx.ts) produces, but for a 100k x 15 workbook this runs in well under a
// second instead of ~12s. See scripts/bench-xlsx.ts and io/fastXlsx.test.ts (differential test
// against readXlsxExceljs).
import { unzipSync } from 'fflate'
import { shiftFormula } from '../engine/shiftFormula'
import { stripFuturePrefixes } from '../engine/xlfn'
import { resolveStructuredRefs, type TableDef } from '../engine/structuredRefs'
import type { BorderSide, CellStyle, CfRule, HyperlinkData, CommentThread, DataValidationRule, DvAllowType, DvErrorStyle, DvOperator, FilterCriteria, NoteData, PageSetup, Range, SheetLayout, CalcMode, SparklineDef } from '../model/types'
import { parseA1 } from '../model/types'
import { parseSheetProtection } from '../model/protection'
import { parsePageSetupXml } from './pageSetupXml'
import type { WorkbookData, DefinedNameData } from '../model/store'
import { DEFAULT_AUTHOR, defaultLayout } from '../model/defaults'
import {
  THEME_COLORS, BORDER_STYLES, applyTint, ptToPx, colWidthToPx, colLetterToIndex, parseRangeStr,
  parseAutoFilterRef, parseAutoFilterColumns, legacyNoteToThread
} from './xlsxShared'
import type { VbaProject } from './xlsxPostProcess'
import {
  decodeXmlEntities,
  parseAttrs,
  extractChild,
  extractSection,
  flattenRichText,
  resolveRelTarget,
  parseRelationships
} from './fastXlsxXml'
import { parseDrawingMlCharts } from './xlsxCharts'

const VBA_REL_TYPE = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
const VBA_SIGNATURE_REL_TYPE_RE = /vbaProjectSignature/

// ---------- built-in number format codes (ECMA-376 18.8.30) ----------
const BUILTIN_NUMFMTS: Record<number, string> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@'
}

function unzipTextFilter(name: string): boolean {
  return (
    name === 'xl/workbook.xml' ||
    name === 'xl/_rels/workbook.xml.rels' ||
    name === 'xl/sharedStrings.xml' ||
    name === 'xl/styles.xml' ||
    name === 'xl/theme/theme1.xml' ||
    /^xl\/worksheets\/[^/]+\.xml$/.test(name) ||
    /^xl\/worksheets\/_rels\/[^/]+\.xml\.rels$/.test(name) ||
    /^xl\/drawings\/[^/]+\.xml$/.test(name) ||
    /^xl\/drawings\/_rels\/[^/]+\.xml\.rels$/.test(name) ||
    /^xl\/charts\/[^/]+\.xml$/.test(name) ||
    /^xl\/tables\/[^/]+\.xml$/.test(name) ||
    // Notes: Excel writes xl/comments1.xml, other producers xl/comments/comment1.xml (found via rels).
    /^xl\/(?:[^/]+\/)?comments?\d*\.xml$/.test(name) ||
    // Raw (non-text) parts, kept only for a .xlsm workbook's VBA project - see parseXlsx below.
    name === 'xl/vbaProject.bin' ||
    name === 'xl/vbaProjectSignature.bin' ||
    name === 'xl/_rels/vbaProject.bin.rels'
  )
}

export function parseXlsx(bytes: Uint8Array): WorkbookData & { vbaProject?: VbaProject } {
  const files = unzipSync(bytes, { filter: (f) => unzipTextFilter(f.name) })
  const dec = new TextDecoder('utf-8')
  const cache = new Map<string, string>()
  const get = (name: string): string | undefined => {
    if (cache.has(name)) return cache.get(name)
    const raw = files[name]
    const text = raw ? dec.decode(raw) : undefined
    cache.set(name, text as string)
    return text
  }

  const workbookXml = get('xl/workbook.xml')
  if (!workbookXml) throw new Error('Invalid .xlsx: missing xl/workbook.xml')

  const relsXml = get('xl/_rels/workbook.xml.rels') ?? ''
  const relMap = new Map<string, string>()
  for (const r of parseRelationships(relsXml)) relMap.set(r.id, r.target)

  const { sheets: sheetDefs, activeTab } = parseWorkbookSheets(workbookXml)
  const sharedStrings = parseSharedStrings(get('xl/sharedStrings.xml'))
  // Theme colors: resolved from theme1.xml when present (falling back to the Office default
  // palette), matching how colorToHex/THEME_COLORS in xlsx.ts treats a theme index + tint.
  const themeColors = parseThemeColors(get('xl/theme/theme1.xml'))
  const xfStyles = parseStyles(get('xl/styles.xml'), themeColors)
  const dxfStyles = parseDxfs(get('xl/styles.xml'), themeColors)

  const sheets: WorkbookData['sheets'] = []
  const tables: TableDef[] = []
  for (const def of sheetDefs) {
    const target = def.rId ? relMap.get(def.rId) : undefined
    const path = target ? resolveRelTarget('xl/workbook.xml', target) : undefined
    const xml = path ? get(path) : undefined

    const layout: Partial<SheetLayout> = { ...defaultLayout() }
    delete (layout as Record<string, unknown>).merges

    if (!xml || !path) {
      layout.merges = []
      layout.notes = {}
      layout.comments = {}
      sheets.push({ name: def.name, cells: {}, layout })
      continue
    }

    const parsed = parseSheetXml(xml, xfStyles, sharedStrings)

    layout.colWidths = parsed.colWidths
    layout.hiddenCols = parsed.hiddenCols
    layout.rowHeights = parsed.rowHeights
    layout.hiddenRows = parsed.hiddenRows
    layout.merges = parsed.merges

    // Notes vs. threaded comments: a legacy note whose text is Excel's own threaded-comment
    // fallback format (see xlsxShared.ts) came from a `comments` thread on save, not a plain note.
    const links = parseHyperlinks(xml, files, dec, path)
    if (Object.keys(links).length) layout.hyperlinks = links
    const sparks = parseSparklines(xml, def.name)
    if (sparks.length) layout.sparklines = sparks
    const notes = parseNotesForSheet(files, dec, path)
    const comments: Record<string, CommentThread> = {}
    for (const [key, note] of Object.entries(notes)) {
      const thread = legacyNoteToThread(note)
      if (thread) {
        comments[key] = thread
        delete notes[key]
      }
    }
    layout.notes = notes
    layout.comments = comments

    if (parsed.showGridlines !== undefined) layout.showGridlines = parsed.showGridlines
    if (parsed.showHeadings !== undefined) layout.showHeadings = parsed.showHeadings
    if (parsed.zoom !== undefined) layout.zoom = parsed.zoom
    if (parsed.frozen) layout.frozen = parsed.frozen
    if (parsed.rowLevels) layout.rowLevels = parsed.rowLevels
    if (parsed.colLevels) layout.colLevels = parsed.colLevels
    if (parsed.collapsedRows) layout.collapsedRows = parsed.collapsedRows
    if (parsed.collapsedCols) layout.collapsedCols = parsed.collapsedCols
    if (parsed.outline) layout.outline = parsed.outline
    const ps = parsePageSetupXml(xml)
    if (ps) layout.pageSetup = ps
    const spAt = xml.lastIndexOf('<sheetProtection')
    if (spAt !== -1) {
      const sp = /^<sheetProtection\b([^>]*?)\/?>/.exec(xml.slice(spAt, spAt + 1500))
      const prot = sp ? parseSheetProtection(parseAttrs(sp[1], 0, sp[1].length)) : undefined
      if (prot) layout.protection = prot
    }
    if (parsed.tabColorAttrs) {
      const c = resolveColor(parseAttrs(parsed.tabColorAttrs, 0, parsed.tabColorAttrs.length), themeColors)
      if (c) layout.tabColor = c
    }
    if (parsed.autoFilterRef) {
      const range = parseSingleOrRange(parsed.autoFilterRef)
      if (range) {
        const columns: Record<number, FilterCriteria> = {}
        for (const [colId, crit] of Object.entries(parseAutoFilterColumns(xml))) {
          columns[Number(colId) + range.c1] = crit
        }
        layout.autoFilter = {
          range,
          columns,
          filteredOutRows: Object.keys(parsed.hiddenRows).map(Number).filter((r) => r > range.r1 && r <= range.r2)
        }
      }
    }

    for (const t of parseTables(files, dec, path, def.name)) {
      tables.push(t.def)
      // A table's filter lives in the table part, not the sheet; keep it as the sheet's AutoFilter.
      if (t.autoFilterXml && !layout.autoFilter) {
        const range = { r1: t.def.r1, c1: t.def.c1, r2: t.def.r2 - t.def.totalsRows, c2: t.def.c2 }
        const columns: Record<number, FilterCriteria> = {}
        for (const [colId, crit] of Object.entries(parseAutoFilterColumns(t.autoFilterXml))) columns[Number(colId) + range.c1] = crit
        layout.autoFilter = { range, columns, filteredOutRows: Object.keys(parsed.hiddenRows).map(Number).filter((r) => r > range.r1 && r <= range.r2) }
      }
    }

    const cfs = parseConditionalFormats(xml, dxfStyles)
    if (cfs.length) layout.conditionalFormats = cfs

    const dvs = parseDataValidations(xml)
    if (dvs.length) layout.dataValidations = dvs
    const charts = parseDrawingMlCharts(files, dec, path, xml, parsed.colWidths, parsed.rowHeights)
    if (charts.length) layout.charts = charts

    sheets.push({ name: def.name, cells: parsed.cells, layout })
  }

  if (tables.length) {
    for (const s of sheets) {
      for (const [key, cell] of Object.entries(s.cells)) {
        if (!cell.raw?.startsWith('=') || cell.raw.indexOf('[') === -1) continue
        cell.raw = '=' + resolveStructuredRefs(cell.raw.slice(1), tables, s.name, Number(key.split(',')[0]))
      }
    }
  }

  sheetDefs.forEach((d, i) => {
    if ((d.state === 'hidden' || d.state === 'veryHidden') && sheets[i]) sheets[i].layout = { ...sheets[i].layout, sheetHidden: d.state }
  })
  const activeSheet = Math.max(0, Math.min(activeTab, sheets.length - 1))
  const names = parseDefinedNames(workbookXml)
  const out: WorkbookData & { vbaProject?: VbaProject } = { sheets, activeSheet, names: names.length ? names : undefined, vbaProject: parseVbaProject(files, relsXml, dec) }
  const wp = /<workbookProtection\b([^>]*?)\/?>/.exec(workbookXml)
  if (wp) {
    const a = parseAttrs(wp[1], 0, wp[1].length)
    const lockStructure = a.lockStructure === '1' || a.lockStructure === 'true', lockWindows = a.lockWindows === '1' || a.lockWindows === 'true'
    if (lockStructure || lockWindows) {
      const password = a.workbookHashValue && a.workbookSaltValue
        ? { hashValue: a.workbookHashValue, saltValue: a.workbookSaltValue, spinCount: parseInt(a.workbookSpinCount ?? '100000', 10) }
        : a.workbookPassword ? { legacy: a.workbookPassword } : undefined
      out.protection = { lockStructure, lockWindows, ...(password ? { password } : {}) }
    }
  }
  const cm = /<calcPr\b[^>]*\bcalcMode="(manual|autoNoTable)"/.exec(workbookXml)
  if (cm) out.calcMode = cm[1] as CalcMode
  const f0 = /<fonts[^>]*>\s*<font[^>]*>([\s\S]*?)<\/font>/.exec(get('xl/styles.xml') ?? '')
  if (f0) {
    const nm = /<name\s+val="([^"]*)"/.exec(f0[1]), sz = /<sz\s+val="([\d.]+)"/.exec(f0[1])
    if (nm && sz && !(nm[1] === 'Calibri' && +sz[1] === 11)) out.defaultFont = { name: nm[1], size: +sz[1] }
  }
  return out
}

/** Excel's x14 sparklineGroups extension -> per-cell sparklines (only ones that read this sheet's own cells). */
function parseSparklines(xml: string, sheetName: string): SparklineDef[] {
  if (!xml.includes('sparklineGroup')) return []
  const out: SparklineDef[] = []
  const unq = (n: string) => n.replace(/^'|'$/g, '').replace(/''/g, "'")
  for (const g of xml.matchAll(/<(?:x14:)?sparklineGroup\b([^>]*)>([\s\S]*?)<\/(?:x14:)?sparklineGroup>/g)) {
    const t = /\btype="(\w+)"/.exec(g[1])?.[1]
    const type: SparklineDef['type'] = t === 'column' ? 'column' : t === 'stacked' ? 'winloss' : 'line'
    const rgb = /<(?:x14:)?colorSeries\b[^>]*\brgb="([0-9A-Fa-f]{8})"/.exec(g[2])?.[1]
    const color = rgb ? '#' + rgb.slice(2).toLowerCase() : '#376092'
    for (const s of g[2].matchAll(/<(?:x14:)?sparkline>\s*<(?:xm:)?f>([^<]*)<\/(?:xm:)?f>\s*<(?:xm:)?sqref>([^<]*)<\/(?:xm:)?sqref>/g)) {
      const f = s[1].replace(/&apos;/g, "'").replace(/&amp;/g, '&')
      const m = /^(?:(.+)!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f)
      const loc = /^\$?([A-Z]+)\$?(\d+)$/.exec(s[2].trim())
      if (!m || !loc || (m[1] && unq(m[1]) !== sheetName)) continue
      const a = parseA1(m[2] + m[3] + (m[4] ? ':' + m[4] + m[5] : ''))
      const at = parseA1(loc[1] + loc[2])
      if (a && at) out.push({ r: at.r1, c: at.c1, type, src: { r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }, color })
    }
  }
  return out
}

function parseTables(files: Record<string, Uint8Array>, dec: TextDecoder, sheetPath: string, sheetName: string): { def: TableDef; autoFilterXml?: string }[] {
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const relsBytes = files[dir + '/_rels/' + (sheetPath.split('/').pop() ?? '') + '.rels']
  if (!relsBytes) return []
  const out: { def: TableDef; autoFilterXml?: string }[] = []
  for (const rel of parseRelationships(dec.decode(relsBytes))) {
    if (!/\/table$/.test(rel.type)) continue
    const bytes = files[resolveRelTarget(sheetPath, rel.target)]
    if (!bytes) continue
    const xml = dec.decode(bytes)
    const head = /<table\b([^>]*)>/.exec(xml)
    if (!head) continue
    const a = parseAttrs(head[1], 0, head[1].length)
    const range = a.ref ? parseSingleOrRange(a.ref) : null
    if (!range) continue
    const columns = [...xml.matchAll(/<tableColumn\b([^>]*?)\/?>/g)].map((m) => decodeXmlEntities(parseAttrs(m[1], 0, m[1].length).name ?? ''))
    out.push({
      def: {
        name: decodeXmlEntities(a.displayName ?? a.name ?? ''), sheet: sheetName, ...range,
        headerRows: a.headerRowCount === '0' ? 0 : 1, totalsRows: a.totalsRowCount ? parseInt(a.totalsRowCount, 10) : 0, columns
      },
      autoFilterXml: /<autoFilter\b[^>]*\/>|<autoFilter\b[\s\S]*?<\/autoFilter>/.exec(xml)?.[0]
    })
  }
  return out
}

// ---------- .xlsm VBA project (kept as opaque bytes so it survives a save back to .xlsm) ----------
function parseVbaProject(files: Record<string, Uint8Array>, workbookRelsXml: string, dec: TextDecoder): VbaProject | undefined {
  const rel = parseRelationships(workbookRelsXml).find((r) => r.type === VBA_REL_TYPE)
  if (!rel) return undefined
  const path = resolveRelTarget('xl/workbook.xml', rel.target)
  const bin = files[path]
  if (!bin) return undefined

  const vba: VbaProject = { bin }
  const sigRelsPath = 'xl/_rels/vbaProject.bin.rels'
  const sigRelsBytes = files[sigRelsPath]
  if (sigRelsBytes) {
    vba.signatureRels = sigRelsBytes
    const sigRelsXml = dec.decode(sigRelsBytes)
    const sigRel = parseRelationships(sigRelsXml).find((r) => VBA_SIGNATURE_REL_TYPE_RE.test(r.type))
    if (sigRel) {
      const sigPath = resolveRelTarget(path, sigRel.target)
      const sigBin = files[sigPath]
      if (sigBin) vba.signature = sigBin
    }
  }
  return vba
}

// ---------- workbook.xml ----------
function parseDefinedNames(xml: string): DefinedNameData[] {
  const section = extractSection(xml, 'definedNames')
  if (!section) return []
  const out: DefinedNameData[] = []
  const re = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(section))) {
    const a = parseAttrs(m[1], 0, m[1].length)
    if (!a.name) continue
    const n: DefinedNameData = { name: a.name, ref: decodeXmlEntities(m[2]).replace(/^=/, '') }
    if (a.localSheetId !== undefined) n.sheet = parseInt(a.localSheetId, 10)
    if (a.hidden === '1' || a.hidden === 'true') n.hidden = true
    out.push(n)
  }
  return out
}

function parseWorkbookSheets(xml: string): { sheets: { name: string; rId: string; state?: string }[]; activeTab: number } {
  const sheets: { name: string; rId: string; state?: string }[] = []
  const section = extractSection(xml, 'sheets')
  if (section) {
    const re = /<sheet\b([^>]*)\/?>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(section))) {
      const a = parseAttrs(m[1], 0, m[1].length)
      if (a.name !== undefined) sheets.push({ name: a.name, rId: a.id ?? '', state: a.state })
    }
  }
  let activeTab = 0
  const bvSection = extractSection(xml, 'bookViews')
  if (bvSection) {
    const m = /<workbookView\b([^>]*)\/?>/.exec(bvSection)
    if (m) {
      const a = parseAttrs(m[1], 0, m[1].length)
      if (a.activeTab) activeTab = parseInt(a.activeTab, 10) || 0
    }
  }
  return { sheets, activeTab }
}

// ---------- sharedStrings.xml ----------
function parseSharedStrings(xml: string | undefined): string[] {
  if (!xml) return []
  const out: string[] = []
  const section = extractSection(xml, 'sst') ?? xml
  let pos = 0
  const n = section.length
  while (true) {
    const idx = section.indexOf('<si', pos)
    if (idx === -1 || idx >= n) break
    const after = idx + 3
    const ch = section.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = idx + 1
      continue
    }
    const gt = section.indexOf('>', after)
    if (gt === -1) break
    if (section.charCodeAt(gt - 1) === 47) {
      out.push('')
      pos = gt + 1
      continue
    }
    const close = section.indexOf('</si>', gt + 1)
    const inner = close === -1 ? section.slice(gt + 1) : section.slice(gt + 1, close)
    out.push(flattenRichText(inner))
    pos = close === -1 ? n : close + 5
  }
  return out
}

// ---------- theme1.xml ----------
// The current exceljs-based reader (readXlsxExceljs / colorToHex in xlsx.ts) resolves a theme
// index through the fixed Office-default palette rather than the workbook's actual theme1.xml, so
// to produce byte-for-byte-equal output (and because exceljs itself always bundles that same
// default theme when writing) we do the same here. `xml` is accepted for interface symmetry with
// a real theme reader, and to leave room for an actual theme1.xml parse if that reader ever changes.
function parseThemeColors(_xml: string | undefined): string[] {
  return THEME_COLORS
}

function resolveColor(attrs: Record<string, string> | undefined, themeColors: string[]): string | undefined {
  if (!attrs) return undefined
  if (attrs.rgb) return '#' + attrs.rgb.slice(-6).toUpperCase()
  if (attrs.theme !== undefined) {
    const idx = parseInt(attrs.theme, 10)
    const base = themeColors[idx] ?? '000000'
    const tint = attrs.tint !== undefined ? parseFloat(attrs.tint) : 0
    return '#' + (tint ? applyTint(base, tint) : base)
  }
  return undefined
}

// ---------- styles.xml ----------
function parseFonts(section: string, themeColors: string[]): (CellStyle | undefined)[] {
  const out: (CellStyle | undefined)[] = []
  let pos = 0
  const n = section.length
  while (true) {
    const idx = section.indexOf('<font', pos)
    if (idx === -1 || idx >= n) break
    const after = idx + 5
    const ch = section.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = idx + 1
      continue
    }
    const gt = section.indexOf('>', after)
    if (gt === -1) break
    if (section.charCodeAt(gt - 1) === 47) {
      out.push(undefined)
      pos = gt + 1
      continue
    }
    const close = section.indexOf('</font>', gt + 1)
    const inner = close === -1 ? section.slice(gt + 1) : section.slice(gt + 1, close)
    pos = close === -1 ? n : close + 7

    const s: CellStyle = {}
    const nameM = /<name val="([^"]*)"/.exec(inner)
    if (nameM) s.fontFamily = decodeXmlEntities(nameM[1])
    const szM = /<sz val="([^"]*)"/.exec(inner)
    if (szM) s.fontSize = parseFloat(szM[1])
    const bM = /<b(?:\s+val="([^"]*)")?\s*\/>/.exec(inner)
    if (bM) s.bold = !(bM[1] === '0' || bM[1] === 'false')
    const iM = /<i(?:\s+val="([^"]*)")?\s*\/>/.exec(inner)
    if (iM) s.italic = !(iM[1] === '0' || iM[1] === 'false')
    const uM = /<u(?:\s+val="([^"]*)")?\s*\/>/.exec(inner)
    if (uM) s.underline = uM[1] === 'double' ? 'double' : 'single'
    const strM = /<strike(?:\s+val="([^"]*)")?\s*\/>/.exec(inner)
    if (strM) s.strike = !(strM[1] === '0' || strM[1] === 'false')
    const colorM = /<color\b([^>]*)\/>/.exec(inner)
    if (colorM) {
      const c = resolveColor(parseAttrs(colorM[1], 0, colorM[1].length), themeColors)
      if (c) s.color = c
    }
    out.push(Object.keys(s).length ? s : undefined)
  }
  return out
}

function parseFills(section: string, themeColors: string[]): (CellStyle | undefined)[] {
  const out: (CellStyle | undefined)[] = []
  let pos = 0
  const n = section.length
  while (true) {
    const idx = section.indexOf('<fill', pos)
    if (idx === -1 || idx >= n) break
    const after = idx + 5
    const ch = section.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = idx + 1
      continue
    }
    const gt = section.indexOf('>', after)
    if (gt === -1) break
    let inner = ''
    let next: number
    if (section.charCodeAt(gt - 1) === 47) {
      next = gt + 1
    } else {
      const close = section.indexOf('</fill>', gt + 1)
      inner = close === -1 ? section.slice(gt + 1) : section.slice(gt + 1, close)
      next = close === -1 ? n : close + 7
    }
    pos = next

    const s: CellStyle = {}
    const patM = /<patternFill\b([^>]*?)(\/)?>/.exec(inner)
    const patternType = patM ? parseAttrs(patM[1], 0, patM[1].length).patternType : undefined
    if (patternType === 'solid') {
      const fgM = /<fgColor\b([^>]*)\/>/.exec(inner)
      if (fgM) {
        const c = resolveColor(parseAttrs(fgM[1], 0, fgM[1].length), themeColors)
        if (c) s.fill = c
      }
    }
    out.push(Object.keys(s).length ? s : undefined)
  }
  return out
}

function parseBorderSide(inner: string, tag: string, themeColors: string[]): BorderSide | undefined {
  const openIdx = inner.indexOf('<' + tag)
  if (openIdx === -1) return undefined
  const after = openIdx + 1 + tag.length
  const ch = inner.charCodeAt(after)
  if (ch !== 62 && ch !== 32 && ch !== 47) return undefined
  const gt = inner.indexOf('>', after)
  if (gt === -1) return undefined
  // <top/> (no attrs) means "no side"; <top style="thin"/> is a self-closing side with no color child.
  const selfClose = inner.charCodeAt(gt - 1) === 47
  const attrsStr = inner.slice(after, selfClose ? gt - 1 : gt)
  const attrs = parseAttrs(attrsStr, 0, attrsStr.length)
  const style = attrs.style
  if (!style) return undefined
  let color: string | undefined
  if (!selfClose) {
    const close = inner.indexOf('</' + tag + '>', gt + 1)
    const body = close === -1 ? '' : inner.slice(gt + 1, close)
    const colorM = /<color\b([^>]*)\/>/.exec(body)
    color = colorM ? resolveColor(parseAttrs(colorM[1], 0, colorM[1].length), themeColors) : undefined
  }
  const validStyle = (BORDER_STYLES as readonly string[]).includes(style) ? (style as (typeof BORDER_STYLES)[number]) : 'thin'
  return { style: validStyle, color }
}

function parseBorders(section: string, themeColors: string[]): (CellStyle | undefined)[] {
  const out: (CellStyle | undefined)[] = []
  let pos = 0
  const n = section.length
  while (true) {
    const idx = section.indexOf('<border', pos)
    if (idx === -1 || idx >= n) break
    const after = idx + 7
    const ch = section.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = idx + 1
      continue
    }
    const gt = section.indexOf('>', after)
    if (gt === -1) break
    let inner = ''
    let next: number
    if (section.charCodeAt(gt - 1) === 47) {
      next = gt + 1
    } else {
      const close = section.indexOf('</border>', gt + 1)
      inner = close === -1 ? section.slice(gt + 1) : section.slice(gt + 1, close)
      next = close === -1 ? n : close + 9
    }
    pos = next

    const top = parseBorderSide(inner, 'top', themeColors)
    const right = parseBorderSide(inner, 'right', themeColors)
    const bottom = parseBorderSide(inner, 'bottom', themeColors)
    const left = parseBorderSide(inner, 'left', themeColors)
    const borders: NonNullable<CellStyle['borders']> = {}
    if (top) borders.top = top
    if (right) borders.right = right
    if (bottom) borders.bottom = bottom
    if (left) borders.left = left
    out.push(Object.keys(borders).length ? { borders } : undefined)
  }
  return out
}

// Differential formats (<dxfs>) referenced by conditional-formatting rules via dxfId.
function parseDxfs(xml: string | undefined, themeColors: string[]): (Partial<CellStyle> | undefined)[] {
  const section = xml ? extractSection(xml, 'dxfs') : undefined
  if (!section) return []
  const out: (Partial<CellStyle> | undefined)[] = []
  const re = /<dxf>([\s\S]*?)<\/dxf>|<dxf\/>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(section))) {
    const inner = m[1] ?? ''
    const s: Partial<CellStyle> = {}
    const font = /<font>[\s\S]*?<\/font>/.exec(inner)
    if (font) {
      const f = parseFonts(font[0], themeColors)[0]
      if (f?.color) s.color = f.color
      if (f?.bold) s.bold = true
      if (f?.italic) s.italic = true
    }
    const fill = /<fill>([\s\S]*?)<\/fill>/.exec(inner)
    if (fill) {
      const c = /<bgColor\b([^>]*?)\/?>/.exec(fill[1]) ?? /<fgColor\b([^>]*?)\/?>/.exec(fill[1])
      const hex = c ? resolveColor(parseAttrs(c[1], 0, c[1].length), themeColors) : undefined
      if (hex) s.fill = hex
    }
    out.push(Object.keys(s).length ? s : undefined)
  }
  return out
}

function parseConditionalFormats(xml: string, dxfs: (Partial<CellStyle> | undefined)[]): CfRule[] {
  if (xml.indexOf('<conditionalFormatting') === -1) return []
  const out: CfRule[] = []
  const blockRe = /<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g
  let b: RegExpExecArray | null
  let pCounter = 1
  while ((b = blockRe.exec(xml))) {
    const ba = parseAttrs(b[1], 0, b[1].length)
    const ranges = (ba.sqref ?? '').split(/\s+/).map(parseSingleOrRange).filter((r): r is Range => r !== null)
    if (!ranges.length) continue
    const ruleRe = /<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g
    let r: RegExpExecArray | null
    while ((r = ruleRe.exec(b[2]))) {
      const a = parseAttrs(r[1], 0, r[1].length)
      const formulae: string[] = []
      const fRe = /<formula>([\s\S]*?)<\/formula>/g
      let f: RegExpExecArray | null
      while ((f = fRe.exec(r[2] ?? ''))) formulae.push(decodeXmlEntities(f[1]))
      const rule: CfRule = {
        id: 'cf_' + Math.random().toString(36).slice(2, 9),
        priority: a.priority ? parseInt(a.priority, 10) : pCounter++,
        ranges,
        stopIfTrue: a.stopIfTrue === '1' || a.stopIfTrue === 'true',
        type: (a.type as CfRule['type']) || 'cellIs',
        format: a.dxfId !== undefined ? dxfs[parseInt(a.dxfId, 10)] : undefined
      }
      if (rule.type === 'cellIs') {
        rule.operator = a.operator as CfRule['operator']
        if (formulae[0] !== undefined) rule.value = formulae[0]
        if (formulae[1] !== undefined) rule.value2 = formulae[1]
      } else if (rule.type === 'expression') {
        rule.formula = formulae[0]
      } else if (rule.type === 'top10') {
        if (a.rank) rule.rank = parseInt(a.rank, 10)
        rule.percent = a.percent === '1'
        rule.bottom = a.bottom === '1'
      } else if (rule.type === 'aboveAverage') {
        rule.above = a.aboveAverage !== '0'
        if (a.equalAverage === '1') rule.equalOrAbove = true
      } else if (rule.type === 'dataBar' || rule.type === 'colorScale' || rule.type === 'iconSet') {
        const colors = [...(r[2] ?? '').matchAll(/<color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/g)].map((m) => '#' + m[1].slice(-6).toUpperCase())
        if (rule.type === 'dataBar') rule.dataBar = { color: colors[0] ?? '#638EC6' }
        else if (rule.type === 'colorScale') {
          rule.colorScale = colors.length >= 3
            ? { minColor: colors[0], midColor: colors[1], maxColor: colors[2] }
            : { minColor: colors[0] ?? '#F8696B', maxColor: colors[1] ?? '#63BE7B' }
        } else {
          const set = /<iconSet\b([^>]*)>/.exec(r[2] ?? '')
          const sa = set ? parseAttrs(set[1], 0, set[1].length) : {}
          rule.iconSet = { name: sa.iconSet ?? '3TrafficLights1', reverse: sa.reverse === '1', showValue: sa.showValue !== '0' }
        }
      }
      out.push(rule)
    }
  }
  return out
}

export function parseDataValidations(xml: string): DataValidationRule[] {
  if (xml.indexOf('<dataValidations') === -1 && xml.indexOf('<dataValidation') === -1) return []
  const out: DataValidationRule[] = []
  const dvRe = /<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g
  let m: RegExpExecArray | null
  while ((m = dvRe.exec(xml))) {
    const a = parseAttrs(m[1], 0, m[1].length)
    const inner = m[2] ?? ''

    const ranges = (a.sqref ?? '')
      .split(/\s+/)
      .map(parseSingleOrRange)
      .filter((r): r is Range => r !== null)
    if (!ranges.length) continue

    const allow = (a.type as DvAllowType) || 'any'
    const rule: DataValidationRule = {
      id: 'dv_' + Math.random().toString(36).slice(2, 9),
      ranges,
      allow
    }

    if (a.operator) {
      rule.operator = a.operator as DvOperator
    } else if (allow !== 'list' && allow !== 'custom' && allow !== 'any') {
      rule.operator = 'between'
    }

    if (a.errorStyle) {
      rule.errorStyle = a.errorStyle as DvErrorStyle
    }

    if (a.allowBlank !== undefined) {
      rule.ignoreBlank = a.allowBlank === '1' || a.allowBlank === 'true'
    }

    // Invert Excel's showDropDown: showDropDown="1" means dropdown is hidden
    if (allow === 'list') {
      rule.inCellDropdown = a.showDropDown !== '1' && a.showDropDown !== 'true'
    }

    if (a.showInputMessage !== undefined) {
      rule.showInputMessage = a.showInputMessage === '1' || a.showInputMessage === 'true'
    }

    if (a.showErrorMessage !== undefined) {
      rule.showErrorAlert = a.showErrorMessage === '1' || a.showErrorMessage === 'true'
    }

    if (a.promptTitle !== undefined) {
      rule.inputTitle = a.promptTitle
    }

    if (a.prompt !== undefined) {
      rule.inputMessage = a.prompt
    }

    if (a.errorTitle !== undefined) {
      rule.errorTitle = a.errorTitle
    }

    if (a.error !== undefined) {
      rule.errorMessage = a.error
    }

    if (inner) {
      const f1 = extractChild(inner, 'formula1', 0, inner.length)
      if (f1) rule.formula1 = decodeXmlEntities(f1.text)
      const f2 = extractChild(inner, 'formula2', 0, inner.length)
      if (f2) rule.formula2 = decodeXmlEntities(f2.text)
    }

    out.push(rule)
  }
  return out
}

function parseStyles(xml: string | undefined, themeColors: string[]): (CellStyle | undefined)[] {
  if (!xml) return [undefined]

  const numFmts = new Map<number, string>()
  const numFmtsSection = extractSection(xml, 'numFmts')
  if (numFmtsSection) {
    const re = /<numFmt\b([^>]*)\/?>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(numFmtsSection))) {
      const a = parseAttrs(m[1], 0, m[1].length)
      if (a.numFmtId && a.formatCode !== undefined) numFmts.set(parseInt(a.numFmtId, 10), a.formatCode)
    }
  }

  const fonts = parseFonts(extractSection(xml, 'fonts') ?? '', themeColors)
  const fills = parseFills(extractSection(xml, 'fills') ?? '', themeColors)
  const borders = parseBorders(extractSection(xml, 'borders') ?? '', themeColors)

  const section = extractSection(xml, 'cellXfs')
  const out: (CellStyle | undefined)[] = []
  if (section) {
    let pos = 0
    const n = section.length
    while (true) {
      const idx = section.indexOf('<xf', pos)
      if (idx === -1 || idx >= n) break
      const after = idx + 3
      const ch = section.charCodeAt(after)
      if (ch !== 62 && ch !== 32 && ch !== 47) {
        pos = idx + 1
        continue
      }
      const gt = section.indexOf('>', after)
      if (gt === -1) break
      const selfClose = section.charCodeAt(gt - 1) === 47
      const attrsStr = section.slice(after, selfClose ? gt - 1 : gt)
      const attrs = parseAttrs(attrsStr, 0, attrsStr.length)
      let inner = ''
      let next: number
      if (selfClose) {
        next = gt + 1
      } else {
        const close = section.indexOf('</xf>', gt + 1)
        inner = close === -1 ? section.slice(gt + 1) : section.slice(gt + 1, close)
        next = close === -1 ? n : close + 5
      }
      pos = next

      const s: CellStyle = {}
      if (attrs.fontId !== undefined) {
        const f = fonts[parseInt(attrs.fontId, 10)]
        if (f) Object.assign(s, f)
      }
      if (attrs.fillId !== undefined) {
        const f = fills[parseInt(attrs.fillId, 10)]
        if (f) Object.assign(s, f)
      }
      if (attrs.borderId !== undefined) {
        const b = borders[parseInt(attrs.borderId, 10)]
        if (b) Object.assign(s, b)
      }
      const numFmtId = attrs.numFmtId !== undefined ? parseInt(attrs.numFmtId, 10) : 0
      const code = numFmts.get(numFmtId) ?? BUILTIN_NUMFMTS[numFmtId]
      if (code && code !== 'General') s.numFmt = decodeXmlEntities(code)

      if (inner) {
        const al = extractChild(inner, 'alignment', 0, inner.length)
        if (al) {
          const a = parseAttrs(al.attrs, 0, al.attrs.length)
          if (a.horizontal) s.hAlign = a.horizontal as CellStyle['hAlign']
          if (a.vertical) s.vAlign = (a.vertical === 'center' ? 'middle' : a.vertical) as CellStyle['vAlign']
          if (a.wrapText === '1') s.wrap = true
          if (a.shrinkToFit === '1') s.shrinkToFit = true
          if (a.indent) {
            const iv = parseInt(a.indent, 10)
            if (iv) s.indent = iv
          }
          if (a.textRotation !== undefined) s.rotation = parseInt(a.textRotation, 10)
        }
        const pr = extractChild(inner, 'protection', 0, inner.length)
        if (pr) {
          const a = parseAttrs(pr.attrs, 0, pr.attrs.length)
          if (a.locked === '0') s.locked = false
          if (a.hidden === '1') s.hidden = true
        }
      }
      out.push(Object.keys(s).length ? s : undefined)
    }
  }
  return out.length ? out : [undefined]
}

// ---------- worksheet xml ----------
interface ParsedSheet {
  cells: WorkbookData['sheets'][number]['cells']
  colWidths: Record<number, number>
  hiddenCols: Record<number, true>
  rowHeights: Record<number, number>
  hiddenRows: Record<number, true>
  merges: Range[]
  autoFilterRef?: string
  tabColorAttrs?: string
  frozen?: { rows: number; cols: number }
  showGridlines?: boolean
  showHeadings?: boolean
  zoom?: number
  rowLevels?: Record<number, number>
  colLevels?: Record<number, number>
  collapsedRows?: Record<number, true>
  collapsedCols?: Record<number, true>
  outline?: { summaryBelow?: boolean; summaryRight?: boolean }
}

type CellsMap = WorkbookData['sheets'][number]['cells']

interface SharedFormulaMaster {
  formula: string
  row: number
  col: number
}

function parseSheetXml(xml: string, xfStyles: (CellStyle | undefined)[], sharedStrings: string[]): ParsedSheet {
  // Object.create(null) avoids prototype-chain bookkeeping V8 otherwise does as this dictionary
  // grows to a million-plus string keys; `cells` is still a plain Record for every consumer (JSON,
  // for-in, Object.keys/entries all work the same on it).
  const cells: CellsMap = Object.create(null) as CellsMap
  const colWidths: Record<number, number> = {}
  const hiddenCols: Record<number, true> = {}
  const rowHeights: Record<number, number> = {}
  const hiddenRows: Record<number, true> = {}
  const merges: Range[] = []
  const mastersBySi = new Map<number, SharedFormulaMaster>()
  const followers: { key: string; row: number; col: number; si: number; style?: CellStyle }[] = []
  const arrayRanges: [number, number, Range][] = []

  // sheetPr / tabColor
  const sheetPr = extractSection(xml, 'sheetPr')
  let tabColorAttrs: string | undefined
  let outline: ParsedSheet['outline']
  if (sheetPr !== undefined) {
    const tc = /<tabColor\b([^>]*)\/>/.exec(sheetPr)
    if (tc) tabColorAttrs = tc[1]
    const op = /<outlinePr\b([^>]*)\/?>/.exec(sheetPr)
    if (op) {
      const a = parseAttrs(op[1], 0, op[1].length)
      outline = {}
      if (a.summaryBelow === '0') outline.summaryBelow = false
      if (a.summaryRight === '0') outline.summaryRight = false
      if (outline.summaryBelow === undefined && outline.summaryRight === undefined) outline = undefined
    }
  }
  // Rows are only scanned for group levels when the sheet declares any (keeps the hot loop untouched otherwise).
  const fmtPr = /<sheetFormatPr\b([^>]*?)\/?>/.exec(xml.slice(0, Math.max(0, xml.indexOf('<sheetData'))))
  const fmtAttrs = fmtPr ? parseAttrs(fmtPr[1], 0, fmtPr[1].length) : {}
  const rowLevels: Record<number, number> | undefined = parseInt(fmtAttrs.outlineLevelRow ?? '0', 10) > 0 ? {} : undefined
  const collapsedRows: Record<number, true> = {}
  const colLevels: Record<number, number> | undefined = parseInt(fmtAttrs.outlineLevelCol ?? '0', 10) > 0 ? {} : undefined
  const collapsedCols: Record<number, true> = {}

  // sheetView
  let showGridlines: boolean | undefined
  let showHeadings: boolean | undefined
  let zoom: number | undefined
  let frozen: { rows: number; cols: number } | undefined
  const sheetViewsSection = extractSection(xml, 'sheetViews')
  if (sheetViewsSection) {
    const svM = /<sheetView\b([^>]*?)(\/?)>/.exec(sheetViewsSection)
    if (svM) {
      const a = parseAttrs(svM[1], 0, svM[1].length)
      if (a.showGridLines !== undefined) showGridlines = a.showGridLines !== '0'
      if (a.showRowColHeaders !== undefined) showHeadings = a.showRowColHeaders !== '0'
      if (a.zoomScale !== undefined) zoom = parseInt(a.zoomScale, 10) || undefined
    }
    const paneM = /<pane\b([^>]*)\/>/.exec(sheetViewsSection)
    if (paneM) {
      const a = parseAttrs(paneM[1], 0, paneM[1].length)
      if (a.state === 'frozen' || a.state === 'frozenSplit') {
        frozen = { rows: a.ySplit ? parseInt(a.ySplit, 10) || 0 : 0, cols: a.xSplit ? parseInt(a.xSplit, 10) || 0 : 0 }
      }
    }
  }

  // cols
  const colsSection = extractSection(xml, 'cols')
  if (colsSection) {
    const re = /<col\b([^>]*)\/?>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(colsSection))) {
      const a = parseAttrs(m[1], 0, m[1].length)
      const min = parseInt(a.min, 10)
      const max = parseInt(a.max, 10)
      if (!Number.isFinite(min) || !Number.isFinite(max)) continue
      const hidden = a.hidden === '1'
      const width = a.width !== undefined ? parseFloat(a.width) : undefined
      // Every explicit width is kept (a width of exactly 9 used to be dropped, copying an exceljs quirk,
      // which silently widened those columns on save).
      for (let i = min; i <= max; i++) {
        if (hidden) hiddenCols[i - 1] = true
        if (width !== undefined) colWidths[i - 1] = colWidthToPx(width)
        if (colLevels && a.outlineLevel && +a.outlineLevel > 0) colLevels[i - 1] = +a.outlineLevel
        if (colLevels && a.collapsed === '1') collapsedCols[i - 1] = true
      }
    }
  }

  // sheetData (the hot path)
  const dataSection = extractSection(xml, 'sheetData')
  if (dataSection) {
    let pos = 0
    const n = dataSection.length
    let rowCounter = 0
    while (true) {
      const rowIdx = dataSection.indexOf('<row', pos)
      if (rowIdx === -1 || rowIdx >= n) break
      const after = rowIdx + 4
      const ch = dataSection.charCodeAt(after)
      if (ch !== 62 && ch !== 32 && ch !== 47) {
        pos = rowIdx + 1
        continue
      }
      const gt = dataSection.indexOf('>', after)
      if (gt === -1) break
      const selfClose = dataSection.charCodeAt(gt - 1) === 47
      const rowAttrsEnd = selfClose ? gt - 1 : gt
      const rAttr = attrVal(dataSection, after, rowAttrsEnd, 'r')
      rowCounter = rAttr !== undefined ? parseInt(rAttr, 10) || rowCounter + 1 : rowCounter + 1
      const r = rowCounter - 1

      const htAttr = attrVal(dataSection, after, rowAttrsEnd, 'ht')
      if (htAttr !== undefined) rowHeights[r] = ptToPx(parseFloat(htAttr))
      if (attrVal(dataSection, after, rowAttrsEnd, 'hidden') === '1') hiddenRows[r] = true
      if (rowLevels) {
        const ol = attrVal(dataSection, after, rowAttrsEnd, 'outlineLevel')
        if (ol && +ol > 0) rowLevels[r] = +ol
        if (attrVal(dataSection, after, rowAttrsEnd, 'collapsed') === '1') collapsedRows[r] = true
      }

      if (selfClose) {
        pos = gt + 1
        continue
      }
      const closeRow = dataSection.indexOf('</row>', gt + 1)
      const rowEnd = closeRow === -1 ? n : closeRow
      parseCellsInRow(dataSection, gt + 1, rowEnd, r, cells, xfStyles, sharedStrings, mastersBySi, followers, arrayRanges)
      pos = closeRow === -1 ? n : closeRow + 6
    }
  }

  // Excel caches an array formula's results (e.g. FILTER's spill) as plain values in the cells it
  // covers; kept, they would block the formula from spilling here and show #SPILL!.
  for (const [ar, ac, rg] of arrayRanges) {
    for (let rr = rg.r1; rr <= rg.r2; rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) {
      if (rr === ar && cc === ac) continue
      const k = rr + ',' + cc
      const cell = cells[k]
      if (cell && !cell.raw?.startsWith('=')) {
        if (cell.style) cells[k] = { style: cell.style }
        else delete cells[k]
      }
    }
  }

  for (const f of followers) {
    const m = mastersBySi.get(f.si)
    const raw = m ? shiftFormula('=' + m.formula, f.row - m.row, f.col - m.col) : undefined
    if (raw !== undefined || f.style) cells[f.key] = { raw, style: f.style }
  }

  // mergeCells. exceljs's `cell.value` on a non-top-left member of a merge returns the master
  // cell's value (it aliases merged cells to their master internally), so readXlsxExceljs reads
  // the same text/formula/number back on every cell in the range, not just the top-left one.
  const mergeSection = extractSection(xml, 'mergeCells')
  if (mergeSection) {
    const re = /<mergeCell ref="([^"]+)"/g
    let m: RegExpExecArray | null
    while ((m = re.exec(mergeSection))) {
      const rg = parseRangeStr(m[1])
      if (rg) merges.push(rg)
    }
  }
  for (const rg of merges) {
    const masterRaw = cells[rg.r1 + ',' + rg.c1]?.raw
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        if (r === rg.r1 && c === rg.c1) continue
        const key = r + ',' + c
        const existing = cells[key]
        if (masterRaw !== undefined || existing) cells[key] = { raw: masterRaw, style: existing?.style }
      }
    }
  }

  // autoFilter
  const autoFilterRef = parseAutoFilterRef(xml)

  return {
    cells,
    colWidths,
    hiddenCols,
    rowHeights,
    hiddenRows,
    merges,
    autoFilterRef,
    tabColorAttrs,
    frozen,
    showGridlines,
    showHeadings,
    zoom,
    rowLevels: rowLevels && Object.keys(rowLevels).length ? rowLevels : undefined,
    colLevels: colLevels && Object.keys(colLevels).length ? colLevels : undefined,
    collapsedRows: Object.keys(collapsedRows).length ? collapsedRows : undefined,
    collapsedCols: Object.keys(collapsedCols).length ? collapsedCols : undefined,
    outline
  }
}

function parseSingleOrRange(ref: string): Range | null {
  const clean = ref.replace(/\$/g, '')
  if (clean.indexOf(':') !== -1) return parseRangeStr(clean)
  const m = /^([A-Z]+)(\d+)$/.exec(clean)
  if (!m) return null
  const c = colLetterToIndex(m[1])
  const r = +m[2] - 1
  return { r1: r, c1: c, r2: r, c2: c }
}

/** Finds `name="value"` within `s[start,end)` without slicing out an attribute substring or
 *  allocating a full attribute map first - the cell/row/formula tags in the hot loop only ever
 *  need one or two named attributes each, and searching the shared backing string with bounds is
 *  cheaper than copying a small substring per tag just to search it. */
function attrVal(s: string, start: number, end: number, name: string): string | undefined {
  // Search only inside [start, end): an unbounded indexOf scans the rest of the sheet whenever the
  // attribute is absent, which made loading quadratic in the number of rows.
  const seg = s.slice(start, end)
  const needle = name + '="'
  let i = seg.indexOf(needle)
  while (i !== -1) {
    const before = i === 0 ? 32 : seg.charCodeAt(i - 1)
    if (before === 32 || before === 9 || before === 10 || before === 13 || before === 47) {
      const vstart = i + needle.length
      const vend = seg.indexOf('"', vstart)
      return vend === -1 ? undefined : seg.slice(vstart, vend)
    }
    i = seg.indexOf(needle, i + 1)
  }
  return undefined
}

/** Text of the first `<tag>...</tag>` (or `''` for a self-closing `<tag/>`) within `s[start,end)`,
 *  bounded the same way as attrVal - used for `<v>`/`<is>`, which are never self-closing with
 *  attributes in practice. */
function tagText(s: string, start: number, end: number, tag: string): string | undefined {
  // Bounded search (see attrVal): a missing <v> must not scan the rest of the sheet.
  const rel = s.slice(start, end).indexOf('<' + tag)
  const idx = rel === -1 ? -1 : start + rel
  if (idx === -1 || idx >= end) return undefined
  const after = idx + tag.length + 1
  const ch = s.charCodeAt(after)
  let contentStart: number
  if (ch === 62 /* '>' */) {
    contentStart = after + 1
  } else if (ch === 47 /* '/' as in <v/> */) {
    return ''
  } else {
    const gt = s.indexOf('>', after)
    if (gt === -1 || gt >= end) return undefined
    if (s.charCodeAt(gt - 1) === 47) return ''
    contentStart = gt + 1
  }
  const relClose = s.slice(contentStart, end).indexOf('</' + tag + '>')
  const closeEnd = relClose === -1 ? end : contentStart + relClose
  return s.slice(contentStart, closeEnd)
}

function colIndexFromRef(ref: string): number {
  let n = 0
  let i = 0
  const len = ref.length
  while (i < len) {
    const cc = ref.charCodeAt(i)
    if (cc < 65 || cc > 90) break
    n = n * 26 + (cc - 64)
    i++
  }
  return n - 1
}

function parseCellsInRow(
  xml: string,
  start: number,
  end: number,
  r: number,
  cells: CellsMap,
  xfStyles: (CellStyle | undefined)[],
  sharedStrings: string[],
  mastersBySi: Map<number, SharedFormulaMaster>,
  followers: { key: string; row: number; col: number; si: number; style?: CellStyle }[],
  arrayRanges: [number, number, Range][]
): void {
  let pos = start
  let colCounter = -1
  while (true) {
    const cIdx = xml.indexOf('<c', pos)
    if (cIdx === -1 || cIdx >= end) break
    const after = cIdx + 2
    const ch = xml.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = cIdx + 1
      continue
    }
    const gt = xml.indexOf('>', after)
    if (gt === -1) break
    const selfClose = xml.charCodeAt(gt - 1) === 47
    const attrsEnd = selfClose ? gt - 1 : gt
    const refAttr = attrVal(xml, after, attrsEnd, 'r')
    const c = refAttr !== undefined ? colIndexFromRef(refAttr) : colCounter + 1
    colCounter = c
    const sAttr = attrVal(xml, after, attrsEnd, 's')
    const style = sAttr !== undefined ? xfStyles[parseInt(sAttr, 10)] : undefined
    const key = r + ',' + c

    if (selfClose) {
      pos = gt + 1
      if (style) cells[key] = { style }
      continue
    }

    const closeC = xml.indexOf('</c>', gt + 1)
    const innerStart = gt + 1
    const innerEnd = closeC === -1 ? end : closeC
    pos = closeC === -1 ? end : closeC + 4

    const fRel = xml.slice(innerStart, innerEnd).indexOf('<f')
    const fIdx = fRel === -1 ? -1 : innerStart + fRel
    if (fIdx !== -1 && fIdx < innerEnd) {
      const fAfter = fIdx + 2
      const fCh = xml.charCodeAt(fAfter)
      const fGt = xml.indexOf('>', fAfter)
      const fSelfClose = fGt !== -1 && fGt < innerEnd && xml.charCodeAt(fGt - 1) === 47
      const fAttrsEnd = fSelfClose ? fGt - 1 : fGt
      const isShared =
        (fCh === 62 || fCh === 32 || fCh === 47) &&
        fGt !== -1 &&
        fGt < innerEnd &&
        attrVal(xml, fAfter, fAttrsEnd, 't') === 'shared'
      let fText: string | undefined
      if (!fSelfClose && fGt !== -1 && fGt < innerEnd) {
        const fClose = xml.indexOf('</f>', fGt + 1)
        const fCloseEnd = fClose === -1 || fClose > innerEnd ? innerEnd : fClose
        fText = xml.slice(fGt + 1, fCloseEnd)
      }
      if (isShared) {
        const siAttr = attrVal(xml, fAfter, fAttrsEnd, 'si')
        const si = siAttr !== undefined ? parseInt(siAttr, 10) : NaN
        if (fText) {
          const formula = stripFuturePrefixes(decodeXmlEntities(fText))
          mastersBySi.set(si, { formula, row: r, col: c })
          cells[key] = { raw: '=' + formula, style }
        } else if (Number.isFinite(si)) {
          followers.push({ key, row: r, col: c, si, style })
        } else if (style) {
          cells[key] = { style }
        }
      } else {
        cells[key] = { raw: '=' + stripFuturePrefixes(decodeXmlEntities(fText ?? '')), style }
        const arrRef = fGt !== -1 && attrVal(xml, fAfter, fAttrsEnd, 't') === 'array' ? attrVal(xml, fAfter, fAttrsEnd, 'ref') : undefined
        const rg = arrRef ? parseSingleOrRange(arrRef) : null
        if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) arrayRanges.push([r, c, rg])
      }
      continue
    }

    const t = attrVal(xml, after, attrsEnd, 't')
    let raw: string | undefined
    if (t === undefined) {
      raw = tagText(xml, innerStart, innerEnd, 'v')
    } else if (t === 's') {
      const v = tagText(xml, innerStart, innerEnd, 'v')
      const idx = v !== undefined ? parseInt(v, 10) : NaN
      raw = Number.isFinite(idx) ? sharedStrings[idx] : undefined
    } else if (t === 'inlineStr') {
      const is = tagText(xml, innerStart, innerEnd, 'is')
      raw = is !== undefined ? flattenRichText(is) : undefined
    } else if (t === 'b') {
      const v = tagText(xml, innerStart, innerEnd, 'v')
      raw = v !== undefined ? (v === '1' ? 'TRUE' : 'FALSE') : undefined
    } else if (t === 'e') {
      raw = tagText(xml, innerStart, innerEnd, 'v')
    } else if (t === 'str') {
      const v = tagText(xml, innerStart, innerEnd, 'v')
      raw = v !== undefined ? decodeXmlEntities(v) : undefined
    } else {
      raw = tagText(xml, innerStart, innerEnd, 'v')
    }

    // Text cells must stay text: without the quote prefix the engine would re-read "3.9.2026" as a
    // date or "00123" as a number, which Excel never does for a cell stored as a string.
    if (raw !== undefined && (t === 's' || t === 'inlineStr' || t === 'str') && needsTextQuote(raw)) raw = "'" + raw
    if (raw !== undefined || style) cells[key] = { raw, style }
  }
}

export function needsTextQuote(s: string): boolean {
  if (s === '') return false
  const c = s.charCodeAt(0)
  if (c === 39 /* ' */ || c === 61 /* = */) return true
  if (/^(true|false)$/i.test(s)) return true
  return /\d/.test(s) && /^[\s\d.,:/\-+$%()€£¥₹eE]*(am|pm|a|p)?\s*$/i.test(s)
}

// ---------- legacy cell notes (xl/comments*.xml) ----------
function parseCommentRuns(inner: string): { bold: boolean; text: string }[] {
  if (inner.indexOf('<r>') === -1 && inner.indexOf('<r ') === -1) {
    const t = extractChild(inner, 't', 0, inner.length)
    return t ? [{ bold: false, text: decodeXmlEntities(t.text) }] : []
  }
  const runs: { bold: boolean; text: string }[] = []
  let pos = 0
  while (true) {
    const i1 = inner.indexOf('<r>', pos)
    const i2 = inner.indexOf('<r ', pos)
    const rIdx = i1 === -1 ? i2 : i2 === -1 ? i1 : Math.min(i1, i2)
    if (rIdx === -1) break
    const close = inner.indexOf('</r>', rIdx)
    const runInner = close === -1 ? inner.slice(rIdx) : inner.slice(rIdx, close)
    const bold = /<b\s*\/?>/.test(runInner)
    const t = extractChild(runInner, 't', 0, runInner.length)
    runs.push({ bold, text: t ? decodeXmlEntities(t.text) : '' })
    pos = close === -1 ? inner.length : close + 4
  }
  return runs
}

function noteFromRuns(runs: { bold: boolean; text: string }[]): NoteData {
  if (runs.length && runs[0].bold && /:\s*$/.test(runs[0].text)) {
    return { author: runs[0].text.replace(/:\s*$/, '').trim(), text: runs.slice(1).map((r) => r.text).join('') }
  }
  return { author: DEFAULT_AUTHOR, text: runs.map((r) => r.text).join('') }
}

function parseCommentsXml(xml: string): Record<string, NoteData> {
  const notes: Record<string, NoteData> = {}
  const section = extractSection(xml, 'commentList')
  if (!section) return notes
  let pos = 0
  const n = section.length
  while (true) {
    const idx = section.indexOf('<comment', pos)
    if (idx === -1 || idx >= n) break
    const after = idx + 8
    const ch = section.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      pos = idx + 1
      continue
    }
    const gt = section.indexOf('>', after)
    if (gt === -1) break
    const selfClose = section.charCodeAt(gt - 1) === 47
    const attrsStr = section.slice(after, selfClose ? gt - 1 : gt)
    const attrs = parseAttrs(attrsStr, 0, attrsStr.length)
    let inner = ''
    let next: number
    if (selfClose) {
      next = gt + 1
    } else {
      const close = section.indexOf('</comment>', gt + 1)
      inner = close === -1 ? section.slice(gt + 1) : section.slice(gt + 1, close)
      next = close === -1 ? n : close + 10
    }
    pos = next
    if (!attrs.ref) continue
    const textSection = extractSection(inner, 'text')
    const runs = textSection !== undefined ? parseCommentRuns(textSection) : []
    const rg = parseSingleOrRange(attrs.ref)
    if (rg) notes[rg.r1 + ',' + rg.c1] = noteFromRuns(runs)
  }
  return notes
}

function parseHyperlinks(xml: string, files: Record<string, Uint8Array>, dec: TextDecoder, sheetPath: string): Record<string, HyperlinkData> {
  const section = extractSection(xml, 'hyperlinks')
  if (!section) return {}
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const file = sheetPath.split('/').pop() ?? ''
  const relsBytes = files[dir + '/_rels/' + file + '.rels']
  const rels = relsBytes ? parseRelationships(dec.decode(relsBytes)) : []
  const out: Record<string, HyperlinkData> = {}
  for (const m of section.matchAll(/<hyperlink\b([^>]*?)\/?>/g)) {
    const a = parseAttrs(m[1], 0, m[1].length)
    const rg = a.ref ? parseSingleOrRange(a.ref) : null
    if (!rg) continue
    const link: HyperlinkData = {}
    if (a.id) { const rel = rels.find((r) => r.id === a.id); if (rel) link.target = rel.target }
    if (a.location) link.location = a.location
    if (a.tooltip) link.tooltip = a.tooltip
    if (!link.target && !link.location) continue
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out[r + ',' + c] = link
  }
  return out
}

function parseNotesForSheet(
  files: Record<string, Uint8Array>,
  dec: TextDecoder,
  sheetPath: string
): Record<string, NoteData> {
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const file = sheetPath.split('/').pop() ?? ''
  const relsPath = dir + '/_rels/' + file + '.rels'
  const relsBytes = files[relsPath]
  if (!relsBytes) return {}
  const relsXml = dec.decode(relsBytes)
  for (const rel of parseRelationships(relsXml)) {
    if (!/comments/i.test(rel.target) && !/\/comments$/.test(rel.type)) continue
    const path = resolveRelTarget(sheetPath, rel.target)
    const bytes = files[path]
    if (!bytes) continue
    return parseCommentsXml(dec.decode(bytes))
  }
  return {}
}

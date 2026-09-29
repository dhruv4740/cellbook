import { addFuturePrefixes } from '../engine/xlfn'
// XLSX read/write via exceljs, mapping to/from the app's WorkbookData model.
import { shiftFormula } from '../engine/shiftFormula'
import ExcelJS from 'exceljs'
import { unzipSync } from 'fflate'
import type { CellStyle, BorderLineStyle, BorderSide, FilterCriteria, SheetLayout, Range, NoteData, CommentThread, CfRule } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, rangeA1, parseA1 } from '../model/types'
import type { WorkbookData } from '../model/store'
import { DEFAULT_AUTHOR, defaultLayout } from '../model/defaults'
import {
  THEME_COLORS, BORDER_STYLES, applyTint, ptToPx, pxToPt, colWidthToPx, pxToColWidth,
  colLetterToIndex, parseRangeStr, parseAutoFilterColumns, threadToLegacyNote, legacyNoteToThread
} from './xlsxShared'
import { extractSection, parseAttrs, parseRelationships, resolveRelTarget } from './fastXlsxXml'
import type { VbaProject } from './xlsxPostProcess'
import { toBytes, fixLegacyDrawingOrder, markDynamicArrays, injectDefinedNames, injectHyperlinks, injectOutline, injectDefaultFont, injectSparklines, injectPageSetup, injectProtection, postProcessPkg } from './xlsxPostProcess'
import { parseDataValidations } from './fastXlsx'
import { injectBulkCells, fastWriteThreshold, BulkRows } from './bulkCells'

const EXCEL_ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'])

// Re-exported for existing `from './xlsx'` call sites (the fast reader/worker import these
// directly from xlsxShared instead, so they don't pull exceljs into their bundle - see fastXlsx.ts).
export {
  THEME_COLORS, BORDER_STYLES, applyTint, ptToPx, pxToPt, colWidthToPx, pxToColWidth,
  colLetterToIndex, parseRangeStr
}
export type { VbaProject }

interface ExcelColor {
  argb?: string
  theme?: number
  tint?: number
  indexed?: number
}

function colorToHex(c: Partial<ExcelColor> | undefined): string | undefined {
  if (!c) return undefined
  if (c.argb) return '#' + c.argb.slice(-6).toUpperCase()
  if (c.theme !== undefined) {
    const base = THEME_COLORS[c.theme] ?? '000000'
    return '#' + (c.tint ? applyTint(base, c.tint) : base)
  }
  return undefined
}

function hexToArgb(hex: string): string {
  const h = hex.replace('#', '').toUpperCase()
  return 'FF' + (h.length === 6 ? h : h.padStart(6, '0'))
}

// ---------- borders ----------
function readBorderSide(b: Partial<ExcelJS.Border> | undefined): BorderSide | undefined {
  if (!b || !b.style) return undefined
  const style = BORDER_STYLES.includes(b.style as BorderLineStyle) ? (b.style as BorderLineStyle) : 'thin'
  const color = colorToHex(b.color as ExcelColor)
  return { style, color }
}

function writeBorderSide(b: BorderSide | undefined): Partial<ExcelJS.Border> | undefined {
  if (!b) return undefined
  return { style: b.style as ExcelJS.BorderStyle, color: { argb: hexToArgb(b.color ?? '#000000') } }
}

// ---------- style mapping ----------
function readStyle(cell: ExcelJS.Cell): CellStyle | undefined {
  const s: CellStyle = {}
  const font = cell.font
  if (font) {
    if (font.name && font.name !== 'Calibri') s.fontFamily = font.name
    else if (font.name) s.fontFamily = font.name
    if (font.size) s.fontSize = font.size
    if (font.bold) s.bold = true
    if (font.italic) s.italic = true
    if (font.underline === true || font.underline === 'single') s.underline = 'single'
    else if (font.underline === 'double') s.underline = 'double'
    if (font.strike) s.strike = true
    const c = colorToHex(font.color as ExcelColor)
    if (c) s.color = c
  }
  const fill = cell.fill
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const c = colorToHex(fill.fgColor as ExcelColor)
    if (c) s.fill = c
  }
  const al = cell.alignment
  if (al) {
    if (al.horizontal) s.hAlign = al.horizontal as CellStyle['hAlign']
    if (al.vertical) s.vAlign = al.vertical as CellStyle['vAlign']
    if (al.wrapText) s.wrap = true
    if (al.shrinkToFit) s.shrinkToFit = true
    if (al.indent) s.indent = al.indent
    if (typeof al.textRotation === 'number') s.rotation = al.textRotation
    else if (al.textRotation === 'vertical') s.rotation = 255
  }
  const numFmt = cell.numFmt
  if (numFmt && numFmt !== 'General') s.numFmt = numFmt
  const bd = cell.border
  if (bd && (bd.top || bd.right || bd.bottom || bd.left)) {
    s.borders = {
      top: readBorderSide(bd.top),
      right: readBorderSide(bd.right),
      bottom: readBorderSide(bd.bottom),
      left: readBorderSide(bd.left)
    }
    for (const k of Object.keys(s.borders) as (keyof typeof s.borders)[]) if (!s.borders[k]) delete s.borders[k]
    if (!Object.keys(s.borders).length) delete s.borders
  }
  const prot = cell.protection
  if (prot) {
    if (prot.locked === false) s.locked = false
    if (prot.hidden) s.hidden = true
  }
  return Object.keys(s).length ? s : undefined
}

function writeStyle(cell: ExcelJS.Cell, style: CellStyle | undefined): void {
  if (!style) return
  const font: Partial<ExcelJS.Font> = {}
  let hasFont = false
  if (style.fontFamily) { font.name = style.fontFamily; hasFont = true }
  if (style.fontSize) { font.size = style.fontSize; hasFont = true }
  if (style.bold) { font.bold = true; hasFont = true }
  if (style.italic) { font.italic = true; hasFont = true }
  if (style.underline) { font.underline = style.underline; hasFont = true }
  if (style.strike) { font.strike = true; hasFont = true }
  if (style.color) { font.color = { argb: hexToArgb(style.color) }; hasFont = true }
  if (hasFont) cell.font = font

  if (style.fill) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(style.fill) } }
  }

  const al: Partial<ExcelJS.Alignment> = {}
  let hasAl = false
  if (style.hAlign && style.hAlign !== 'general') { al.horizontal = style.hAlign; hasAl = true }
  if (style.vAlign) { al.vertical = style.vAlign; hasAl = true }
  if (style.wrap) { al.wrapText = true; hasAl = true }
  if (style.shrinkToFit) { al.shrinkToFit = true; hasAl = true }
  if (style.indent) { al.indent = style.indent; hasAl = true }
  if (style.rotation) { al.textRotation = style.rotation === 255 ? 'vertical' : style.rotation; hasAl = true }
  if (hasAl) cell.alignment = al

  if (style.numFmt) cell.numFmt = style.numFmt

  if (style.borders) {
    cell.border = {
      top: writeBorderSide(style.borders.top),
      right: writeBorderSide(style.borders.right),
      bottom: writeBorderSide(style.borders.bottom),
      left: writeBorderSide(style.borders.left)
    }
  }

  if (style.locked === false || style.hidden) {
    cell.protection = { locked: style.locked !== false ? true : false, hidden: !!style.hidden }
  }
}

// ---------- date handling ----------
const EXCEL_EPOCH = Date.UTC(1899, 11, 30)
const MS_PER_DAY = 24 * 60 * 60 * 1000

function dateToSerial(d: Date): number {
  return (d.getTime() - EXCEL_EPOCH) / MS_PER_DAY
}
function serialToDate(n: number): Date {
  return new Date(EXCEL_EPOCH + n * MS_PER_DAY)
}

function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false
  // Strip quoted/escaped/bracketed sections, then look for date/time tokens.
  const stripped = fmt.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')
  return /[ymdhs]/i.test(stripped) && !/^general$/i.test(fmt)
}

// exceljs's object model doesn't expose a worksheet's raw XML (and has no filterColumn support at
// all), so for the AutoFilter-criteria read path below we separately unzip the same bytes and map
// each sheet name to its raw xl/worksheets/sheetN.xml text, the same way fastXlsx.ts does.
function mapSheetXmlByName(data: Uint8Array): Map<string, string> {
  const map = new Map<string, string>()
  const files = unzipSync(data, {
    filter: (f) => f.name === 'xl/workbook.xml' || f.name === 'xl/_rels/workbook.xml.rels' || /^xl\/worksheets\/[^/]+\.xml$/.test(f.name)
  })
  const dec = new TextDecoder('utf-8')
  const workbookXml = files['xl/workbook.xml'] ? dec.decode(files['xl/workbook.xml']) : undefined
  if (!workbookXml) return map
  const relsXml = files['xl/_rels/workbook.xml.rels'] ? dec.decode(files['xl/_rels/workbook.xml.rels']) : ''
  const relMap = new Map<string, string>()
  for (const r of parseRelationships(relsXml)) relMap.set(r.id, r.target)
  const section = extractSection(workbookXml, 'sheets')
  if (!section) return map
  const re = /<sheet\b([^>]*)\/?>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(section))) {
    const a = parseAttrs(m[1], 0, m[1].length)
    if (a.name === undefined || !a.id) continue
    const target = relMap.get(a.id)
    if (!target) continue
    const path = resolveRelTarget('xl/workbook.xml', target)
    const xml = files[path]
    if (xml) map.set(a.name, dec.decode(xml))
  }
  return map
}

// ---------- read (exceljs, kept for the differential test against the fast reader) ----------
export async function readXlsxExceljs(data: Uint8Array): Promise<WorkbookData> {
  const wb = new ExcelJS.Workbook()
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await wb.xlsx.load(data as any)
  const sheetXmlByName = mapSheetXmlByName(data)

  const sheets: WorkbookData['sheets'] = []
  wb.eachSheet((ws) => {
    const cells: WorkbookData['sheets'][number]['cells'] = {}
    const notes: Record<string, NoteData> = {}
    ws.eachRow({ includeEmpty: true }, (row, rowNum) => {
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const r = rowNum - 1
        const c = colNum - 1
        const key = r + ',' + c
        const style = readStyle(cell)
        const noteVal = cell.note as string | ExcelJS.Comment | undefined
        if (noteVal) notes[key] = readNote(noteVal)
        let raw: string | undefined
        const v = cell.value
        if (v === null || v === undefined) raw = undefined
        else if (typeof v === 'object' && 'formula' in (v as object)) {
          const fv = v as ExcelJS.CellFormulaValue
          raw = '=' + fv.formula
        } else if (typeof v === 'object' && 'sharedFormula' in (v as object)) {
          // Followers of a shared formula only name their master cell: shift the master's formula here.
          const fv = v as ExcelJS.CellSharedFormulaValue
          const master = ws.getCell(fv.sharedFormula)
          const mf = (master.value as ExcelJS.CellFormulaValue | null)?.formula
          raw = fv.formula
            ? '=' + fv.formula
            : mf
              ? shiftFormula('=' + mf, rowNum - Number(master.row), colNum - Number(master.col))
              : undefined
        } else if (typeof v === 'object' && 'richText' in (v as object)) {
          raw = (v as ExcelJS.CellRichTextValue).richText.map((rt) => rt.text).join('')
        } else if (typeof v === 'object' && 'text' in (v as object) && 'hyperlink' in (v as object)) {
          raw = (v as ExcelJS.CellHyperlinkValue).text
        } else if (v instanceof Date) {
          raw = String(dateToSerial(v))
          if (!style || !style.numFmt) {
            // ensure round-trip still reads as date-ish General fallback: keep m/d/yyyy
          }
        } else if (typeof v === 'boolean') {
          raw = v ? 'TRUE' : 'FALSE'
        } else if (typeof v === 'object' && 'error' in (v as object)) {
          raw = (v as ExcelJS.CellErrorValue).error
        } else {
          raw = String(v)
        }
        if (raw === undefined && !style) return
        cells[key] = { raw, style }
      })
    })

    const layout: Partial<SheetLayout> = { ...defaultLayout() }
    delete (layout as Record<string, unknown>).merges

    // column widths / hidden
    const colWidths: Record<number, number> = {}
    const hiddenCols: Record<number, true> = {}
    const colCount = Math.max(ws.columnCount || 0, ws.columns?.length ?? 0)
    for (let i = 1; i <= colCount; i++) {
      const col = ws.getColumn(i)
      if (col.hidden) hiddenCols[i - 1] = true
      if (col.width !== undefined) colWidths[i - 1] = colWidthToPx(col.width)
    }
    layout.colWidths = colWidths
    layout.hiddenCols = hiddenCols

    const rowHeights: Record<number, number> = {}
    const hiddenRows: Record<number, true> = {}
    ws.eachRow({ includeEmpty: true }, (row, rowNum) => {
      if (row.hidden) hiddenRows[rowNum - 1] = true
      if (row.height !== undefined) rowHeights[rowNum - 1] = ptToPx(row.height)
    })
    layout.rowHeights = rowHeights
    layout.hiddenRows = hiddenRows

    // merges
    const merges: Range[] = []
    const modelMerges = (ws.model as ExcelJS.WorksheetModel | undefined)?.merges ?? []
    for (const m of modelMerges) {
      const rg = parseRangeStr(m)
      if (rg) merges.push(rg)
    }
    layout.merges = merges

    // Notes vs. threaded comments: a legacy note whose text is Excel's own threaded-comment
    // fallback format (see xlsxShared.ts) came from a `comments` thread on save, not a plain note.
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

    // AutoFilter: exceljs's object model only round-trips the filtered range, not per-column
    // criteria (there's no filterColumn support at all), so those are parsed from the sheet's raw
    // XML instead, via the same parser the fast reader uses.
    const af = ws.autoFilter as ExcelJS.AutoFilter | string | undefined
    let range: Range | undefined
    if (typeof af === 'string') {
      range = parseRangeStr(af) ?? undefined
    } else if (af && typeof af === 'object') {
      const from = typeof af.from === 'object' ? af.from : parseA1CellRef(af.from)
      const to = typeof af.to === 'object' ? af.to : parseA1CellRef(af.to)
      if (from && to) range = { r1: from.row - 1, c1: from.column - 1, r2: to.row - 1, c2: to.column - 1 }
    }
    if (range) {
      const sheetXml = sheetXmlByName.get(ws.name)
      const columns: Record<number, FilterCriteria> = {}
      if (sheetXml) {
        for (const [colId, crit] of Object.entries(parseAutoFilterColumns(sheetXml))) {
          columns[Number(colId) + range.c1] = crit
        }
      }
      const r1 = range.r1
      const r2 = range.r2
      layout.autoFilter = {
        range,
        columns,
        filteredOutRows: Object.keys(hiddenRows).map(Number).filter((r) => r > r1 && r <= r2)
      }
    }

    // frozen panes
    const view = ws.views && ws.views[0]
    if (view && (view as ExcelJS.WorksheetViewFrozen).state === 'frozen') {
      const fv = view as ExcelJS.WorksheetViewFrozen
      layout.frozen = { rows: fv.ySplit ?? 0, cols: fv.xSplit ?? 0 }
    }
    if (view) {
      if (view.showGridLines !== undefined) layout.showGridlines = view.showGridLines
      if (view.showRowColHeaders !== undefined) layout.showHeadings = view.showRowColHeaders
      if (view.zoomScale !== undefined) layout.zoom = view.zoomScale
    }
    if (ws.state === 'hidden' || ws.state === 'veryHidden') layout.sheetHidden = ws.state
    const tabColor = (ws.properties as ExcelJS.WorksheetProperties | undefined)?.tabColor as ExcelColor | undefined
    if (tabColor) {
      const c = colorToHex(tabColor)
      if (c) layout.tabColor = c
    }

    const rawCfs = (ws as any).conditionalFormattings || []
    if (rawCfs.length > 0) {
      const cfs: CfRule[] = []
      let pCounter = 1
      for (const cf of rawCfs) {
        const refStr = cf.ref || ''
        const ranges = refStr
          .split(/\s+/)
          .map((r: string) => parseA1(r))
          .filter((r: Range | null): r is Range => r !== null)
        if (!ranges.length) continue

        for (const rule of cf.rules || []) {
          const id = 'cf_' + Math.random().toString(36).slice(2, 9)
          const priority = rule.priority ?? pCounter++
          const stopIfTrue = Boolean(rule.stopIfTrue)
          const format: Partial<CellStyle> = {}
          if (rule.style?.fill?.bgColor) {
            format.fill = colorToHex(rule.style.fill.bgColor)
          } else if (rule.style?.fill?.fgColor) {
            format.fill = colorToHex(rule.style.fill.fgColor)
          }
          if (rule.style?.font) {
            if (rule.style.font.color) format.color = colorToHex(rule.style.font.color)
            if (rule.style.font.bold) format.bold = true
            if (rule.style.font.italic) format.italic = true
          }

          let cfRule: CfRule = {
            id,
            priority,
            ranges,
            stopIfTrue,
            type: (rule.type as any) || 'cellIs',
            format: Object.keys(format).length > 0 ? format : undefined
          }

          if (rule.type === 'cellIs') {
            cfRule.operator = rule.operator
            if (Array.isArray(rule.formulae)) {
              cfRule.value = String(rule.formulae[0] ?? '')
              if (rule.formulae[1] !== undefined) cfRule.value2 = String(rule.formulae[1])
            }
          } else if (rule.type === 'expression') {
            cfRule.formula = Array.isArray(rule.formulae) ? rule.formulae[0] : rule.formula
          } else if (rule.type === 'top10') {
            cfRule.rank = rule.rank
            cfRule.percent = rule.percent
            cfRule.bottom = rule.bottom
          } else if (rule.type === 'aboveAverage') {
            cfRule.above = rule.aboveAverage ?? true
            cfRule.equalOrAbove = rule.equalAverage
          }

          cfs.push(cfRule)
        }
      }
      if (cfs.length > 0) layout.conditionalFormats = cfs
    }

    const sheetXml = sheetXmlByName.get(ws.name)
    if (sheetXml) {
      const dvs = parseDataValidations(sheetXml)
      if (dvs.length > 0) layout.dataValidations = dvs
    }

    sheets.push({ name: ws.name, cells, layout })
  })

  let activeSheet = 0
  const activeTab = (wb.views && wb.views[0] && (wb.views[0] as ExcelJS.WorkbookView).activeTab) ?? 0
  activeSheet = Math.max(0, Math.min(activeTab, sheets.length - 1))

  return { sheets, activeSheet }
}

function parseA1CellRef(s: string | undefined): { row: number; column: number } | null {
  if (!s) return null
  const m = /^([A-Z]+)(\d+)$/.exec(s.split(':')[0])
  if (!m) return null
  return { row: +m[2], column: colLetterToIndex(m[1]) + 1 }
}

// ---------- notes (exceljs "note" = the classic cell comment box, not a threaded Comment) ----------
function readNote(v: string | ExcelJS.Comment): NoteData {
  if (typeof v === 'string') return { text: v, author: DEFAULT_AUTHOR }
  const runs = v.texts ?? []
  if (runs.length && runs[0].font?.bold && /:\s*$/.test(runs[0].text)) {
    return { author: runs[0].text.replace(/:\s*$/, '').trim(), text: runs.slice(1).map((r) => r.text).join('') }
  }
  return { author: DEFAULT_AUTHOR, text: runs.map((r) => r.text).join('') }
}
function writeNote(note: NoteData): ExcelJS.Comment {
  return { texts: [{ font: { bold: true }, text: note.author + ':\n' }, { text: note.text }] }
}

// ---------- write ----------
/** `vba`, when given, is re-injected into the written package after exceljs writes it (exceljs
 *  itself has no concept of a VBA project and always drops one) - pass it only when saving to
 *  .xlsm; saving to .xlsx must drop macros, same as real Excel. */
export async function writeXlsx(data: WorkbookData, vba?: VbaProject): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Excel Clone'
  wb.created = new Date()
  wb.calcProperties.fullCalcOnLoad = true

  // Big sheets: their plain cells are written as XML by bulkCells.ts instead of through exceljs objects.
  const bulk: (BulkRows | undefined)[] = []
  const threshold = fastWriteThreshold()
  data.sheets.forEach((sd, i) => {
    const ws = wb.addWorksheet(sd.name || 'Sheet' + (i + 1))
    const layout = { ...defaultLayout(), ...sd.layout }

    // Cells that need exceljs objects (notes, comments, merged areas) stay on the exceljs path.
    let special: Set<string> | undefined
    if (Object.keys(sd.cells).length >= threshold) {
      special = new Set<string>([...Object.keys(layout.notes ?? {}), ...Object.keys(layout.comments ?? {})])
      let mergedCells = 0
      for (const m of layout.merges ?? []) mergedCells += (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1)
      if (mergedCells > 50000) special = undefined
      else for (const m of layout.merges ?? []) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) special.add(`${r},${c}`)
    }
    const rows: BulkRows | undefined = special ? new Map() : undefined
    bulk[i] = rows

    for (const key in sd.cells) {
      const cell = sd.cells[key]
      const comma = key.indexOf(',')
      const r = +key.slice(0, comma), c = +key.slice(comma + 1)
      if (rows && !special!.has(key)) {
        let row = rows.get(r)
        if (!row) rows.set(r, (row = []))
        row.push({ c, cell })
        continue
      }
      const target = ws.getCell(r + 1, c + 1)
      if (cell.raw !== undefined && cell.raw !== '') {
        if (cell.raw.startsWith("'")) {
          target.value = cell.raw.slice(1) // quote-prefixed text
        } else if (cell.raw.startsWith('=')) {
          // Store the computed result too, as Excel does: Protected View, previews and other tools
          // read it without recalculating.
          const v = cell.value
          // Only Excel's own error codes may be cached; e.g. HyperFormula's "#ERROR!" makes Excel refuse the file.
          const err = v && typeof v === 'object' ? (v as { error: string }).error : undefined
          const result = v === null || v === undefined || (err !== undefined && !EXCEL_ERRORS.has(err)) ? undefined : err !== undefined ? { error: err } : v
          const formula = addFuturePrefixes(cell.raw.slice(1))
          const spill = cell.spill ? { shareType: 'array', ref: cell.spill } : {}
          target.value = (result === undefined ? { formula, ...spill } : { formula, result, ...spill }) as ExcelJS.CellValue
        } else if (cell.style?.numFmt && isDateFormat(cell.style.numFmt) && /^-?\d+(\.\d+)?$/.test(cell.raw)) {
          target.value = serialToDate(Number(cell.raw))
        } else if (cell.raw === 'TRUE' || cell.raw === 'FALSE') {
          target.value = cell.raw === 'TRUE'
        } else if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(cell.raw)) {
          target.value = Number(cell.raw)
        } else {
          target.value = cell.raw
        }
      }
      if (cell.style) writeStyle(target, cell.style)
    }

    // Comments win over a plain note at the same cell: a threaded comment is saved as a legacy
    // note in Excel's own fallback format (xlsxShared.ts), so a cell can't show both anyway.
    const noteEntries: Record<string, NoteData> = { ...layout.notes }
    for (const [key, thread] of Object.entries(layout.comments ?? {})) {
      noteEntries[key] = threadToLegacyNote(thread)
    }
    for (const [key, note] of Object.entries(noteEntries)) {
      const [r, c] = key.split(',').map(Number)
      ws.getCell(r + 1, c + 1).note = writeNote(note)
    }

    for (const [c, px] of Object.entries(layout.colWidths)) {
      ws.getColumn(+c + 1).width = pxToColWidth(px)
    }
    for (const c of Object.keys(layout.hiddenCols)) {
      ws.getColumn(+c + 1).hidden = true
    }
    for (const [r, px] of Object.entries(layout.rowHeights)) {
      ws.getRow(+r + 1).height = pxToPt(px)
    }
    for (const r of Object.keys(layout.hiddenRows)) {
      const row = ws.getRow(+r + 1)
      row.hidden = true
      // exceljs only serializes a row that has cells or an explicit height; give hidden-but-empty rows one
      // so the hidden flag round-trips through the file.
      if (row.height === undefined) row.height = pxToPt(layout.rowHeights[+r] ?? DEFAULT_ROW_HEIGHT)
    }
    // Group levels. ExcelJS derives `collapsed` wrongly, so injectOutline() rewrites those flags afterwards.
    for (const [r, n] of Object.entries(layout.rowLevels ?? {})) {
      const row = ws.getRow(+r + 1)
      row.outlineLevel = n
      if (!row.hasValues && row.height === undefined) row.height = pxToPt(layout.rowHeights[+r] ?? DEFAULT_ROW_HEIGHT)
    }
    for (const [c, n] of Object.entries(layout.colLevels ?? {})) {
      const col = ws.getColumn(+c + 1)
      col.outlineLevel = n
      if (col.width === undefined) col.width = pxToColWidth(layout.colWidths[+c] ?? DEFAULT_COL_WIDTH)
    }
    ws.properties.outlineLevelRow = Math.max(0, ...Object.values(layout.rowLevels ?? {}))
    ws.properties.outlineLevelCol = Math.max(0, ...Object.values(layout.colLevels ?? {}))
    for (const r of Object.keys(layout.collapsedRows ?? {})) {
      const row = ws.getRow(+r + 1)
      if (!row.hasValues && row.height === undefined) row.height = pxToPt(layout.rowHeights[+r] ?? DEFAULT_ROW_HEIGHT)
    }
    for (const c of Object.keys(layout.collapsedCols ?? {})) {
      const col = ws.getColumn(+c + 1)
      if (col.width === undefined) col.width = pxToColWidth(layout.colWidths[+c] ?? DEFAULT_COL_WIDTH)
    }
    if (layout.outline) ws.properties.outlineProperties = { summaryBelow: layout.outline.summaryBelow !== false, summaryRight: layout.outline.summaryRight !== false }
    for (const m of layout.merges) {
      ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1)
    }
    if (layout.sheetHidden) ws.state = layout.sheetHidden
    if (layout.tabColor) {
      ws.properties.tabColor = { argb: hexToArgb(layout.tabColor) }
    }
    if (layout.autoFilter) {
      const { range } = layout.autoFilter
      ws.autoFilter = {
        from: { row: range.r1 + 1, column: range.c1 + 1 },
        to: { row: range.r2 + 1, column: range.c2 + 1 }
      }
    }
    if (layout.conditionalFormats && layout.conditionalFormats.length > 0) {
      for (const cfRule of layout.conditionalFormats) {
        const ref = cfRule.ranges.map(rangeA1).join(' ')
        const style: any = {}
        if (cfRule.format) {
          if (cfRule.format.fill) {
            style.fill = {
              type: 'pattern',
              pattern: 'solid',
              bgColor: { argb: hexToArgb(cfRule.format.fill) }
            }
          }
          if (cfRule.format.color || cfRule.format.bold || cfRule.format.italic) {
            style.font = {}
            if (cfRule.format.color) style.font.color = { argb: hexToArgb(cfRule.format.color) }
            if (cfRule.format.bold) style.font.bold = true
            if (cfRule.format.italic) style.font.italic = true
          }
        }
        const ruleObj: any = {
          priority: cfRule.priority,
          stopIfTrue: cfRule.stopIfTrue
        }
        if (cfRule.type === 'cellIs') {
          ruleObj.type = 'cellIs'
          ruleObj.operator = cfRule.operator
          ruleObj.formulae = cfRule.value2 !== undefined ? [cfRule.value, cfRule.value2] : [cfRule.value]
          ruleObj.style = style
        } else if (cfRule.type === 'expression') {
          ruleObj.type = 'expression'
          ruleObj.formulae = [cfRule.formula?.replace(/^=/, '')]
          ruleObj.style = style
        } else if (cfRule.type === 'top10') {
          ruleObj.type = 'top10'
          ruleObj.rank = cfRule.rank || 10
          ruleObj.percent = cfRule.percent
          ruleObj.bottom = cfRule.bottom
          ruleObj.style = style
        } else if (cfRule.type === 'aboveAverage') {
          ruleObj.type = 'aboveAverage'
          ruleObj.aboveAverage = cfRule.above ?? true
          ruleObj.equalAverage = cfRule.equalOrAbove
          ruleObj.style = style
        } else if (cfRule.type === 'dataBar') {
          // ExcelJS needs the min/max thresholds; without them it throws while saving.
          ruleObj.type = 'dataBar'
          ruleObj.cfvo = [{ type: 'min' }, { type: 'max' }]
          ruleObj.color = { argb: hexToArgb(cfRule.dataBar?.color ?? '#638EC6') }
        } else if (cfRule.type === 'colorScale') {
          const cs = cfRule.colorScale ?? { minColor: '#F8696B', maxColor: '#63BE7B' }
          ruleObj.type = 'colorScale'
          ruleObj.cfvo = cs.midColor ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }]
          ruleObj.color = (cs.midColor ? [cs.minColor, cs.midColor, cs.maxColor] : [cs.minColor, cs.maxColor]).map((c) => ({ argb: hexToArgb(c) }))
        } else if (cfRule.type === 'iconSet') {
          const name = cfRule.iconSet?.name ?? '3TrafficLights1'
          const n = parseInt(name, 10) || 3
          ruleObj.type = 'iconSet'
          ruleObj.iconSet = name
          ruleObj.reverse = cfRule.iconSet?.reverse
          ruleObj.showValue = cfRule.iconSet?.showValue
          ruleObj.cfvo = Array.from({ length: n }, (_, i) => ({ type: 'percent', value: Math.round((i * 100) / n) }))
        } else {
          ruleObj.type = cfRule.type
          ruleObj.style = style
        }
        ;(ws as any).addConditionalFormatting({
          ref,
          rules: [ruleObj]
        })
      }
    }

    const view: Partial<ExcelJS.WorksheetView> = {
      showGridLines: layout.showGridlines,
      showRowColHeaders: layout.showHeadings,
      zoomScale: layout.zoom,
      zoomScaleNormal: layout.zoom
    }
    if (layout.frozen.rows > 0 || layout.frozen.cols > 0) {
      Object.assign(view, {
        state: 'frozen',
        xSplit: layout.frozen.cols,
        ySplit: layout.frozen.rows
      })
    }
    ws.views = [view]
  })

  // Excel refuses a workbook whose sheets are all hidden, or whose active tab is hidden.
  const isVisible = (i: number) => !data.sheets[i]?.layout?.sheetHidden
  if (!data.sheets.some((_, i) => isVisible(i))) wb.worksheets[0].state = 'visible'
  let activeTab = data.activeSheet ?? 0
  if (!isVisible(activeTab)) activeTab = Math.max(0, data.sheets.findIndex((_, i) => isVisible(i)))
  wb.views = [{ x: 0, y: 0, width: 25000, height: 15000, firstSheet: 0, activeTab, visibility: 'visible' }]

  const anyBulk = bulk.some(Boolean)
  const raw = new Uint8Array(await wb.xlsx.writeBuffer(anyBulk ? ({ zip: { compression: 'STORE' } } as never) : undefined) as ArrayBuffer)
  const buf = anyBulk ? injectBulkCells(raw, bulk) : raw

  // AutoFilter per-column criteria and the VBA project part both need editing the raw package -
  // exceljs's object model can express neither. Skip the unzip/rezip when there's nothing to add.
  const sheetFilters = data.sheets.map((sd) => sd.layout?.autoFilter)
  const sheetValidations = data.sheets.map((sd) => sd.layout?.dataValidations)
  const sheetCharts = data.sheets.map((sd) => sd.layout?.charts)
  const sheetNames = data.sheets.map((sd, i) => sd.name || `Sheet${i + 1}`)
  const sheetCells = data.sheets.map((sd) => sd.cells)
  const sheetLayouts = data.sheets.map((sd) => sd.layout)

  const hasCriteria = sheetFilters.some((af) => af && Object.keys(af.columns).length)
  const hasValidations = sheetValidations.some((dv) => dv && dv.length > 0)
  const hasCharts = sheetCharts.some((c) => c && c.length > 0)

  const processed = hasCriteria || hasValidations || hasCharts || vba
    ? postProcessPkg(buf, sheetFilters, vba, sheetValidations, sheetCharts, sheetNames, sheetCells, sheetLayouts)
    : buf
  const linked = markDynamicArrays(fixLegacyDrawingOrder(injectHyperlinks(processed, data.sheets.map((sd) => sd.layout?.hyperlinks))))
  const named = data.names?.length ? injectDefinedNames(linked, data.names) : linked
  return toBytes(injectDefaultFont(injectSparklines(injectPageSetup(injectOutline(injectProtection(named, data.sheets.map((sd) => sd.layout?.protection), data.protection, data.calcMode), data.sheets.map((sd) => sd.layout)), data.sheets.map((sd) => sd.layout?.pageSetup)), data.sheets.map((sd) => sd.layout?.sparklines), sheetNames), data.defaultFont))
}

/** Same as writeXlsx, but big workbooks are built in a Web Worker (the UI stays responsive while saving). */
export async function writeXlsxOffThread(data: WorkbookData, vba?: VbaProject): Promise<Uint8Array> {
  let cells = 0
  for (const s of data.sheets) cells += Object.keys(s.cells).length
  if (cells < 20000 || typeof Worker === 'undefined' || typeof window === 'undefined') return writeXlsx(data, vba)
  try {
    return await new Promise<Uint8Array>((resolve, reject) => {
      const worker = new Worker(new URL('./xlsxWrite.worker.ts', import.meta.url), { type: 'module' })
      worker.onmessage = (ev: MessageEvent<{ ok: true; bytes: Uint8Array } | { ok: false; error: string }>): void => {
        worker.terminate()
        if (ev.data.ok) resolve(ev.data.bytes)
        else reject(new Error(ev.data.error))
      }
      worker.onerror = (ev: ErrorEvent): void => {
        worker.terminate()
        reject(ev.error instanceof Error ? ev.error : new Error(ev.message || 'xlsx write worker failed'))
      }
      worker.postMessage({ data, vba })
    })
  } catch {
    return writeXlsx(data, vba) // the worker could not run here: build it on this thread instead
  }
}

// ---------- read (fast path) ----------
// Parsing is the hot path for big files, so the real work lives in fastXlsx.ts (fflate unzip + a
// hand-written XML scanner, no exceljs/DOMParser). In a browser/Electron renderer it runs inside a
// Web Worker so the UI thread never blocks; in Node (tests, tooling) it falls back to parsing inline.
export async function readXlsx(data: Uint8Array): Promise<WorkbookData & { vbaProject?: VbaProject }> {
  if (typeof Worker !== 'undefined' && typeof window !== 'undefined') {
    return new Promise<WorkbookData & { vbaProject?: VbaProject }>((resolve, reject) => {
      let worker: Worker
      try {
        worker = new Worker(new URL('./xlsx.worker.ts', import.meta.url), { type: 'module' })
      } catch {
        // Worker construction unsupported in this host (e.g. certain test runners) - fall back inline.
        void parseXlsxInline(data).then(resolve, reject)
        return
      }
      const ab = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
      worker.onmessage = (ev: MessageEvent<{ ok: true; data: WorkbookData & { vbaProject?: VbaProject } } | { ok: false; error: string }>): void => {
        worker.terminate()
        if (ev.data.ok) resolve(ev.data.data)
        else reject(new Error(ev.data.error))
      }
      worker.onerror = (ev: ErrorEvent): void => {
        worker.terminate()
        reject(ev.error instanceof Error ? ev.error : new Error(ev.message || 'xlsx worker failed'))
      }
      worker.postMessage(ab, [ab])
    })
  }
  return parseXlsxInline(data)
}

async function parseXlsxInline(data: Uint8Array): Promise<WorkbookData & { vbaProject?: VbaProject }> {
  const { parseXlsx } = await import('./fastXlsx')
  return parseXlsx(data)
}

// XLSX read/write via exceljs, mapping to/from the app's WorkbookData model.
import { shiftFormula } from '../engine/shiftFormula'
import ExcelJS from 'exceljs'
import type { CellStyle, BorderLineStyle, BorderSide, SheetLayout, Range, CfRule } from '../model/types'
import { DEFAULT_ROW_HEIGHT, rangeA1, parseA1 } from '../model/types'
import type { WorkbookData } from '../model/store'
import { defaultLayout } from '../model/store'

// ---------- column width / row height conversions ----------
// Excel character-width <-> pixel conversion for the default 11pt font (Maximum Digit Width ~7px).
export function colWidthToPx(w: number): number {
  return Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7)
}
export function pxToColWidth(px: number): number {
  if (px <= 0) return 0
  return Math.round(((px * 256) / 7 - Math.trunc(128 / 7)) / 256 * 100) / 100
}
// Row height: points <-> pixels at 96 dpi.
export function ptToPx(pt: number): number {
  return Math.round((pt * 4) / 3)
}
export function pxToPt(px: number): number {
  return Math.round(((px * 3) / 4) * 100) / 100
}

// ---------- Office default theme colors (Office 2013+ theme) ----------
const THEME_COLORS = [
  'FFFFFF', // 0 bg1 (light1)
  '000000', // 1 tx1 (dark1)
  'E7E6E6', // 2 bg2 (light2)
  '44546A', // 3 tx2 (dark2)
  '4472C4', // 4 accent1
  'ED7D31', // 5 accent2
  'A5A5A5', // 6 accent3
  'FFC000', // 7 accent4
  '5B9BD5', // 8 accent5
  '70AD47' // 9 accent6
]

function applyTint(hex: string, tint: number): string {
  // Excel tint: positive lightens toward white, negative darkens toward black.
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  const adj = (c: number): number => {
    if (tint < 0) return Math.round(c * (1 + tint))
    return Math.round(c * (1 - tint) + 255 * tint)
  }
  const toHex = (n: number): string => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')
  return (toHex(adj(r)) + toHex(adj(g)) + toHex(adj(b))).toUpperCase()
}

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
const BORDER_STYLES: BorderLineStyle[] = [
  'thin', 'medium', 'thick', 'double', 'dotted', 'dashed', 'hair',
  'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'
]

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

// ---------- read ----------
export async function readXlsx(data: Uint8Array): Promise<WorkbookData> {
  const wb = new ExcelJS.Workbook()
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await wb.xlsx.load(data as any)

  const sheets: WorkbookData['sheets'] = []
  wb.eachSheet((ws) => {
    const cells: WorkbookData['sheets'][number]['cells'] = {}
    ws.eachRow({ includeEmpty: true }, (row, rowNum) => {
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const r = rowNum - 1
        const c = colNum - 1
        const key = r + ',' + c
        const style = readStyle(cell)
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
      if (col.width !== undefined && col.isCustomWidth) colWidths[i - 1] = colWidthToPx(col.width)
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

    sheets.push({ name: ws.name, cells, layout })
  })

  let activeSheet = 0
  const activeTab = (wb.views && wb.views[0] && (wb.views[0] as ExcelJS.WorkbookView).activeTab) ?? 0
  activeSheet = Math.max(0, Math.min(activeTab, sheets.length - 1))

  return { sheets, activeSheet }
}

function parseRangeStr(s: string): Range | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(s)
  if (!m) return null
  return { r1: +m[2] - 1, c1: colLetterToIndex(m[1]), r2: +m[4] - 1, c2: colLetterToIndex(m[3]) }
}
function colLetterToIndex(s: string): number {
  let n = 0
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

// ---------- write ----------
export async function writeXlsx(data: WorkbookData): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Excel Clone'
  wb.created = new Date()
  wb.calcProperties.fullCalcOnLoad = true

  data.sheets.forEach((sd, i) => {
    const ws = wb.addWorksheet(sd.name || 'Sheet' + (i + 1))
    const layout = { ...defaultLayout(), ...sd.layout }

    for (const [key, cell] of Object.entries(sd.cells)) {
      const [r, c] = key.split(',').map(Number)
      const target = ws.getCell(r + 1, c + 1)
      if (cell.raw !== undefined && cell.raw !== '') {
        if (cell.raw.startsWith('=')) {
          target.value = { formula: cell.raw.slice(1) }
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
    for (const m of layout.merges) {
      ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1)
    }
    if (layout.tabColor) {
      ws.properties.tabColor = { argb: hexToArgb(layout.tabColor) }
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

  wb.views = [{ x: 0, y: 0, width: 25000, height: 15000, firstSheet: 0, activeTab: data.activeSheet ?? 0, visibility: 'visible' }]

  const buf = await wb.xlsx.writeBuffer()
  return new Uint8Array(buf as ArrayBuffer)
}

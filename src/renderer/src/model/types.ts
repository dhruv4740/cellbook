// Core data model shared by every part of the app. Row/column indexes are 0-based.

export type BorderLineStyle =
  | 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'

export interface BorderSide {
  style: BorderLineStyle
  color?: string // '#RRGGBB'
}

export interface CellStyle {
  fontFamily?: string // default 'Aptos Narrow'
  fontSize?: number // points, default 11
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'double'
  strike?: boolean
  color?: string // font color '#RRGGBB'
  fill?: string // background '#RRGGBB'
  hAlign?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
  vAlign?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed'
  wrap?: boolean
  shrinkToFit?: boolean
  indent?: number
  rotation?: number // degrees -90..90, or 255 for vertical text
  numFmt?: string // Excel number format code, e.g. 'General', '0.00', '#,##0', 'm/d/yyyy'
  borders?: { top?: BorderSide; right?: BorderSide; bottom?: BorderSide; left?: BorderSide }
  locked?: boolean
  hidden?: boolean
}

export interface Range {
  r1: number
  c1: number
  r2: number
  c2: number
}

export interface CellRef {
  r: number
  c: number
}

export interface Selection {
  ranges: Range[] // last range is the one being extended
  active: CellRef // the white cell inside the selection
}

export interface SheetLayout {
  colWidths: Record<number, number> // px, default DEFAULT_COL_WIDTH
  rowHeights: Record<number, number> // px, default DEFAULT_ROW_HEIGHT
  hiddenRows: Record<number, true>
  hiddenCols: Record<number, true>
  merges: Range[]
  frozen: { rows: number; cols: number }
  showGridlines: boolean
  showHeadings: boolean
  zoom: number // 10..400 percent
  tabColor?: string
  dataValidations?: DataValidationRule[]
}

export type DvAllowType =
  | 'any'
  | 'whole'
  | 'decimal'
  | 'list'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom'

export type DvOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual'

export type DvErrorStyle = 'stop' | 'warning' | 'information'

export interface DataValidationRule {
  id: string
  ranges: Range[]
  allow: DvAllowType
  operator?: DvOperator
  formula1?: string
  formula2?: string
  ignoreBlank?: boolean
  inCellDropdown?: boolean
  showInputMessage?: boolean
  inputTitle?: string
  inputMessage?: string
  showErrorAlert?: boolean
  errorStyle?: DvErrorStyle
  errorTitle?: string
  errorMessage?: string
}

export type CellValue = number | string | boolean | null | CellError

export interface CellError {
  error: '#DIV/0!' | '#N/A' | '#NAME?' | '#NULL!' | '#NUM!' | '#REF!' | '#VALUE!' | '#SPILL!' | '#CALC!' | '#CYCLE!' | '#ERROR!'
}

export const DEFAULT_COL_WIDTH = 64 // Excel 8.43 chars at 100% zoom
export const DEFAULT_ROW_HEIGHT = 20 // 15pt
export const MAX_ROWS = 1048576
export const MAX_COLS = 16384
export const DEFAULT_FONT = 'Aptos Narrow'
export const DEFAULT_FONT_SIZE = 11

export function isError(v: CellValue): v is CellError {
  return typeof v === 'object' && v !== null && 'error' in v
}

export function cellKey(r: number, c: number): string {
  return r + ',' + c
}

export function colName(c: number): string {
  let s = ''
  c += 1
  while (c > 0) {
    const m = (c - 1) % 26
    s = String.fromCharCode(65 + m) + s
    c = Math.floor((c - 1) / 26)
  }
  return s
}

export function colIndex(name: string): number {
  let n = 0
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

export function a1(r: number, c: number): string {
  return colName(c) + (r + 1)
}

export function rangeA1(rg: Range): string {
  const n = normRange(rg)
  return n.r1 === n.r2 && n.c1 === n.c2 ? a1(n.r1, n.c1) : a1(n.r1, n.c1) + ':' + a1(n.r2, n.c2)
}

/** Parses 'B3', 'A1:C5', 'A:A', '3:3' (no sheet prefix). Returns null if invalid. */
export function parseA1(ref: string): Range | null {
  const s = ref.trim().toUpperCase().replace(/\$/g, '')
  let m = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(s)
  if (m) {
    const r1 = +m[2] - 1, c1 = colIndex(m[1])
    const r2 = m[4] ? +m[4] - 1 : r1, c2 = m[3] ? colIndex(m[3]) : c1
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS || c1 >= MAX_COLS || c2 >= MAX_COLS) return null
    return normRange({ r1, c1, r2, c2 })
  }
  m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s)
  if (m) return normRange({ r1: 0, r2: MAX_ROWS - 1, c1: colIndex(m[1]), c2: colIndex(m[2]) })
  m = /^(\d{1,7}):(\d{1,7})$/.exec(s)
  if (m) return normRange({ r1: +m[1] - 1, r2: +m[2] - 1, c1: 0, c2: MAX_COLS - 1 })
  return null
}

export function normRange(rg: Range): Range {
  return {
    r1: Math.min(rg.r1, rg.r2), r2: Math.max(rg.r1, rg.r2),
    c1: Math.min(rg.c1, rg.c2), c2: Math.max(rg.c1, rg.c2)
  }
}

export function inRange(rg: Range, r: number, c: number): boolean {
  const n = normRange(rg)
  return r >= n.r1 && r <= n.r2 && c >= n.c1 && c <= n.c2
}

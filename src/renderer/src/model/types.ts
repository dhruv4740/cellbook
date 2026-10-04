// Core data model shared by every part of the app. Row/column indexes are 0-based.

export type BorderLineStyle =
  | 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'

export interface BorderSide {
  style: BorderLineStyle
  color?: string // '#RRGGBB'
}

export interface CellStyle {
  fontFamily?: string // default 'Calibri' (Excel 2021's body font)
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

/** One column's checked-values criteria within an AutoFilter. `selected` lists the display
 *  values (as text) whose rows stay visible; the sentinel BLANK_TOKEN (engine/filter.ts) stands
 *  for empty cells. Columns with no entry here are unfiltered. */
export interface FilterCriteria {
  selected: string[]
}

export interface AutoFilter {
  range: Range // includes the header row at range.r1
  columns: Record<number, FilterCriteria> // key: absolute column index
  /** Rows currently hidden by this filter (recomputed on every criteria change), so toggling the
   *  filter off only unhides rows the filter itself hid, not ones hidden by Home > Hide Rows. */
  filteredOutRows: number[]
}

export interface NoteData {
  text: string
  author: string
  /** Review > Show/Hide Note: the note stays on screen instead of only on hover. */
  visible?: boolean
}

export interface CommentReply {
  id: string
  author: string
  text: string
  time: string // ISO
}

export interface CommentThread {
  author: string
  text: string
  time: string // ISO
  resolved?: boolean
  replies: CommentReply[]
}

/** Page Layout / Page Setup, saved in the file (worksheet pageSetup, pageMargins, printOptions, headerFooter, breaks). */
export interface PageSetup {
  landscape?: boolean
  /** Excel paper size code (9 = A4, 1 = Letter, 5 = Legal...), see model/pageSetup.ts. */
  paperSize?: number
  /** Inches; missing sides use the Normal preset. */
  margins?: { left?: number; right?: number; top?: number; bottom?: number; header?: number; footer?: number }
  /** Adjust to N% of normal size (10-400) when fitToPage is off. */
  scale?: number
  /** Fit to N pages wide by N tall; 0 = automatic. Only used when fitToPage is set. */
  fitToPage?: boolean
  fitToWidth?: number
  fitToHeight?: number
  gridlines?: boolean
  headings?: boolean
  blackAndWhite?: boolean
  draft?: boolean
  centerH?: boolean
  centerV?: boolean
  pageOrder?: 'overThenDown'
  firstPageNumber?: number
  comments?: 'atEnd' | 'asDisplayed'
  errors?: 'blank' | 'dash' | 'NA'
  /** Header & footer strings in Excel's code language (&L &C &R &P &N ...). */
  hf?: {
    oddHeader?: string; oddFooter?: string; evenHeader?: string; evenFooter?: string; firstHeader?: string; firstFooter?: string
    differentOddEven?: boolean; differentFirst?: boolean; scaleWithDoc?: boolean; alignWithMargins?: boolean
  }
  /** Manual page breaks: 0-based index of the first row / column of the page that follows the break. */
  rowBreaks?: number[]
  colBreaks?: number[]
}

/** Excel's password hash: SHA-512 (agile) for files saved by Excel 2013+, or the legacy 16-bit hash. */
export interface PasswordHash { legacy?: string; hashValue?: string; saltValue?: string; spinCount?: number }

export type ProtectAllowKey =
  | 'selectLockedCells' | 'selectUnlockedCells' | 'formatCells' | 'formatColumns' | 'formatRows'
  | 'insertColumns' | 'insertRows' | 'insertHyperlinks' | 'deleteColumns' | 'deleteRows'
  | 'sort' | 'autoFilter' | 'pivotTables' | 'objects' | 'scenarios'

/** Review > Protect Sheet. `allow[k]` is true when users may do k on the protected sheet. */
export interface SheetProtection {
  password?: PasswordHash
  allow: Record<ProtectAllowKey, boolean>
}

/** `ref` is a space-separated list of A1 ranges (Excel's sqref), e.g. "A1:B4 D6". */
export interface EditRange { title: string; ref: string }

export type CalcMode ='auto' | 'autoNoTable' | 'manual'

export interface WorkbookProtection {
  password?: PasswordHash
  lockStructure: boolean
  lockWindows: boolean
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
  /** Home > Format > Hide Sheet (Excel's sheet state; 'veryHidden' is only settable from VBA). */
  sheetHidden?: 'hidden' | 'veryHidden'
  /** Page Layout settings used by Print / PDF and saved to the file. */
  pageSetup?: PageSetup
  /** Data > Group: group depth per row / column (absent = not grouped), see model/outline.ts. */
  rowLevels?: Record<number, number>
  colLevels?: Record<number, number>
  /** Summary lines whose group was collapsed by the user (remembered so nested groups re-expand correctly). */
  collapsedRows?: Record<number, true>
  collapsedCols?: Record<number, true>
  /** Data > Outline settings; both default to true (summary below / to the right of the detail). */
  outline?: { summaryBelow?: boolean; summaryRight?: boolean }
  /** Set while the sheet is protected (Review > Protect Sheet). */
  protection?: SheetProtection
  /** Review > Allow Edit Ranges: cells that stay editable on a protected sheet. Kept while the sheet is unprotected. */
  editRanges?: EditRange[]
  /** PivotTables whose output starts on this sheet (kept for Refresh; saved to .xlsx as values). */
  pivots?: import('../data/pivot').PivotDef[]
  autoFilter?: AutoFilter
  /** Data > Advanced Filter (in place): the rows it hid, so Data > Clear can show them again. */
  advFilter?: { hidden: number[] }
  notes: Record<string, NoteData> // key 'r,c'
  comments: Record<string, CommentThread> // key 'r,c'
  /** Cell hyperlinks, key 'r,c': an external target (URL, file, mailto) or an in-workbook location. */
  hyperlinks?: Record<string, HyperlinkData>
  /** Insert > Sparklines: one mini chart per cell, drawn from a range on the same sheet. */
  sparklines?: SparklineDef[]
  /** Insert > Pictures: floating images (bottom of the list is drawn first). */
  pictures?: PictureDef[]
  /** Insert > Shapes / Text Box. */
  shapes?: ShapeDef[]
  /** Excel tables (ListObjects) kept so a save writes them back; formulas already use plain ranges. */
  tables?: SheetTable[]
  /** Data > What-If Analysis > Scenario Manager. */
  scenarios?: ScenarioDef[]
  /** Floating chart objects on this sheet. Optional so existing layouts still typecheck. */
  charts?: ChartSpec[]
  conditionalFormats?: CfRule[]
  dataValidations?: DataValidationRule[]
}

// ---------- Charts ----------
export type ChartType = 'bar' | 'line' | 'pie' | 'scatter'
export type ChartSubtype =
  | 'clusteredColumn' | 'stackedColumn' | 'percentStackedColumn'
  | 'clusteredBar' | 'stackedBar' | 'percentStackedBar'
  | 'line' | 'stackedLine' | 'lineMarkers'
  | 'area' | 'stackedArea'
  | 'pie' | 'doughnut'
  | 'scatter' | 'scatterSmooth'

export type LegendPosition = 'right' | 'top' | 'bottom' | 'left' | 'none'

export interface ChartElementsVisible {
  title: boolean
  legend: boolean
  gridlines: boolean
  dataLabels: boolean
  axes: boolean
  axisTitles: boolean
}

/** Page Layout > Arrange state shared by charts, pictures and shapes. */
export interface ArrangeProps {
  /** Stacking order across all floating objects on the sheet (higher = in front); absent = on top of any that have one. */
  z?: number
  /** Hidden via the Selection Pane. */
  hidden?: boolean
  /** Objects sharing a group id move and select together. */
  group?: string
  /** Degrees clockwise (pictures and shapes only). */
  rotation?: number
  flipH?: boolean
  flipV?: boolean
}

/** Anchors the chart's top-left corner to a cell, offset by pixels within that cell (at 100% zoom). */
export interface ChartAnchor {
  r: number
  c: number
  dx: number
  dy: number
}

export interface ChartSpec extends ArrangeProps {
  id: string
  type: ChartType
  subtype: ChartSubtype
  /** Source range the chart plots (Excel's "current region" when inserted, or an explicit Select Data range). */
  dataRange: Range
  /** true: each column of dataRange is a series (categories run down rows). false: each row is a series. */
  seriesInColumns: boolean
  title: string
  anchor: ChartAnchor
  width: number // px at 100% zoom
  height: number
  legendPosition: LegendPosition
  axisTitles?: { x?: string; y?: string }
  elements: ChartElementsVisible
  /** Index into the Change Colors gallery (theme color rotation variants). */
  colorSetId?: number
  /** Index into the Chart Styles gallery. */
  styleId?: number
  /** Series indices hidden via the Chart Filters funnel (Select Data still lists them). */
  hiddenSeriesIdx?: number[]
  /** Treat the first column (or row, when series are in rows) as category labels even when it is numeric, e.g. a timeline. */
  firstIsCategories?: boolean
}

export type CellValue = number | string | boolean | null | CellError

export interface CellError {
  error: '#DIV/0!' | '#N/A' | '#NAME?' | '#NULL!' | '#NUM!' | '#REF!' | '#VALUE!' | '#SPILL!' | '#CALC!' | '#CYCLE!' | '#ERROR!'
}

export const DEFAULT_COL_WIDTH = 64 // Excel 8.43 chars at 100% zoom
export const DEFAULT_ROW_HEIGHT = 20 // 15pt
export const MAX_ROWS = 1048576
export const MAX_COLS = 16384
export const DEFAULT_FONT = 'Calibri'
export const DEFAULT_FONT_SIZE = 11
/** The open workbook's Normal-style font (set on load); cells without their own font draw with it. */
export const normalFont = { name: DEFAULT_FONT, size: DEFAULT_FONT_SIZE }
/** Office 2013-2022 theme colors ("Office" theme), used by every color picker and chart default. */
export const THEME_COLORS = {
  dark1: '#000000', light1: '#FFFFFF', dark2: '#44546A', light2: '#E7E6E6',
  accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47'
}

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

// ---------- Conditional formatting ----------
export type CfRuleType =
  | 'cellIs'
  | 'textContains'
  | 'dateOccurring'
  | 'duplicateValues'
  | 'uniqueValues'
  | 'top10'
  | 'aboveAverage'
  | 'expression'
  | 'dataBar'
  | 'colorScale'
  | 'iconSet'

export type CfOperator =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual'

export type CfDateOccurring =
  | 'yesterday'
  | 'today'
  | 'tomorrow'
  | 'inTheLast7Days'
  | 'lastWeek'
  | 'thisWeek'
  | 'nextWeek'
  | 'lastMonth'
  | 'thisMonth'
  | 'nextMonth'

export interface CfRule {
  id: string
  ranges: Range[]
  priority: number
  stopIfTrue?: boolean
  type: CfRuleType
  operator?: CfOperator
  value?: string
  value2?: string
  text?: string
  date?: CfDateOccurring
  percent?: boolean
  bottom?: boolean
  rank?: number
  above?: boolean
  equalOrAbove?: boolean
  formula?: string
  format?: Partial<CellStyle>
  dataBar?: {
    color: string
    min?: number
    max?: number
    gradient?: boolean
    showValue?: boolean
  }
  colorScale?: {
    minColor: string
    midColor?: string
    maxColor: string
  }
  iconSet?: {
    name: string
    reverse?: boolean
    showValue?: boolean
  }
}

// ---------- Data validation ----------
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

export interface HyperlinkData { target?: string; location?: string; tooltip?: string }

export interface ScenarioDef {
  name: string
  /** The changing cells and the value (raw text) each takes in this scenario. */
  cells: { r: number; c: number }[]
  values: string[]
  comment?: string
}

export interface SheetTable {
  name: string
  r1: number
  c1: number
  r2: number
  c2: number
  headerRows: number
  totalsRows: number
  /** Style name such as TableStyleMedium2 (absent = no style). */
  style?: string
  rowStripes?: boolean
  colStripes?: boolean
  firstCol?: boolean
  lastCol?: boolean
  /** The table shows filter buttons. */
  filter?: boolean
}

export type ShapeKind = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'diamond' | 'rightArrow' | 'line' | 'arrow' | 'textbox'

export interface ShapeDef extends ArrangeProps {
  id: string
  name: string
  kind: ShapeKind
  anchor: { r: number; c: number; dx: number; dy: number }
  width: number
  height: number
  /** Fill colour (#rrggbb); absent = no fill. */
  fill?: string
  /** Outline colour; absent = no outline. */
  line?: string
  lineWidth?: number
  text?: string
  color?: string
  fontSize?: number
  bold?: boolean
  italic?: boolean
  align?: 'l' | 'ctr' | 'r'
  /** WordArt-style text effects. */
  fontFamily?: string
  /** Text outline colour (#rrggbb). */
  textOutline?: string
  textShadow?: boolean
}

export interface PictureDef extends ArrangeProps {
  id: string
  name: string
  /** `data:image/...;base64,...` */
  src: string
  /** Top-left corner: a cell plus a pixel offset into it. */
  anchor: { r: number; c: number; dx: number; dy: number }
  /** Size in pixels at 100% zoom. */
  width: number
  height: number
  /** Alt text. */
  alt?: string
}

export interface SparklineDef {
  r: number
  c: number
  type: 'line' | 'column' | 'winloss'
  /** Source cells (same sheet); a row or a column. */
  src: { r1: number; c1: number; r2: number; c2: number }
  color: string
}

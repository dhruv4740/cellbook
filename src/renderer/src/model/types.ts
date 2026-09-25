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
  autoFilter?: AutoFilter
  notes: Record<string, NoteData> // key 'r,c'
  comments: Record<string, CommentThread> // key 'r,c'
  /** Floating chart objects on this sheet. Optional so existing layouts still typecheck. */
  charts?: ChartSpec[]
  /** Floating pictures/shapes/text boxes (and preserved-verbatim objects we can't model) on this
   *  sheet. Optional so existing layouts still typecheck. See DrawingObject. Not yet wired into
   *  the UI/IO layers - see the handoff note in this session's final report. */
  drawings?: DrawingObject[]
  conditionalFormats?: CfRule[]
  dataValidations?: DataValidationRule[]
  /** Page Layout tab / File > Print / Page Setup dialog. Optional so existing layouts still typecheck. */
  pageSetup?: PageSetup
  /** Excel Tables (ListObjects) - Home > Format as Table / Insert > Table / Ctrl+T. Optional so
   *  existing layouts still typecheck. */
  tables?: Table[]
}

// ---------- Tables (ListObjects) ----------
/** Total Row per-column aggregate, matching the dropdown Excel shows under each column of the
 *  Total Row. Written to the cell as `=SUBTOTAL(<code>,<range>)`; 'custom' means the user typed
 *  something else via "More Functions" and `customFormula` holds it verbatim (with '='). */
export type TotalsRowFunction =
  | 'none' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'sum' | 'stdDev' | 'var' | 'custom'

/** SUBTOTAL function_num for each aggregate, using the 100+ range so hidden (filtered) rows are
 *  ignored - matching what Excel's own Total Row writes. */
export const TOTALS_SUBTOTAL_CODE: Record<Exclude<TotalsRowFunction, 'none' | 'custom'>, number> = {
  average: 101, count: 103, countNums: 102, max: 104, min: 105, sum: 109, stdDev: 107, var: 110
}
export const TOTALS_ROW_LABELS: Record<TotalsRowFunction, string> = {
  none: 'None', average: 'Average', count: 'Count', countNums: 'Count Numbers', max: 'Max', min: 'Min',
  sum: 'Sum', stdDev: 'StdDev', var: 'Var', custom: 'More Functions...'
}

export interface TableColumn {
  name: string
  totalsRowFunction?: TotalsRowFunction // absent/'none' = no aggregate (blank or label text)
  /** Only meaningful when totalsRowFunction === 'custom': the raw A1 formula (with leading '='). */
  totalsRowFormula?: string
  /** Plain text label typed into the Total Row cell instead of a function (e.g. "Total"). */
  totalsRowLabel?: string
  /** A calculated column: the A1 formula (with leading '=') as entered in the column's first data
   *  row; other rows hold the same formula shifted down (engine/shiftFormula.ts), like Excel's
   *  calculated columns. Undefined for an ordinary (non-formula) column. */
  calculatedFormula?: string
}

export interface TableStyleOptions {
  headerRow: boolean
  totalRow: boolean
  bandedRows: boolean
  firstColumn: boolean
  lastColumn: boolean
  bandedColumns: boolean
  filterButton: boolean
}

export function defaultTableStyleOptions(): TableStyleOptions {
  return { headerRow: true, totalRow: false, bandedRows: true, firstColumn: false, lastColumn: false, bandedColumns: false, filterButton: true }
}

export interface Table {
  id: string
  name: string
  /** Full extent including the header row (if headerRow) and the Total Row (if totalRow). */
  range: Range
  styleName: string // e.g. 'TableStyleMedium2', or 'TableStyleNone'
  options: TableStyleOptions
  columns: TableColumn[] // one per column of range, in range order
}

/** The data rows of a table: excludes the header row (if any) and the Total Row (if any). */
export function tableDataRange(t: Table): Range {
  const r1 = t.options.headerRow ? t.range.r1 + 1 : t.range.r1
  const r2 = t.options.totalRow ? t.range.r2 - 1 : t.range.r2
  return { r1: Math.min(r1, r2), c1: t.range.c1, r2, c2: t.range.c2 }
}
export function tableHeaderRow(t: Table): number | null {
  return t.options.headerRow ? t.range.r1 : null
}
export function tableTotalRow(t: Table): number | null {
  return t.options.totalRow ? t.range.r2 : null
}

// ---------- Printing (Page Layout tab, Page Setup dialog, File > Print) ----------
export type PaperSize = 'letter' | 'legal' | 'tabloid' | 'executive' | 'a3' | 'a4' | 'a5'

/** Paper dimensions in inches (portrait), matching Windows/Excel's built-in sizes. */
export const PAPER_SIZES: Record<PaperSize, { label: string; w: number; h: number; ooxml: number }> = {
  letter: { label: 'Letter', w: 8.5, h: 11, ooxml: 1 },
  legal: { label: 'Legal', w: 8.5, h: 14, ooxml: 5 },
  executive: { label: 'Executive', w: 7.25, h: 10.5, ooxml: 7 },
  a3: { label: 'A3', w: 11.69, h: 16.54, ooxml: 8 },
  a4: { label: 'A4', w: 8.27, h: 11.69, ooxml: 9 },
  a5: { label: 'A5', w: 5.83, h: 8.27, ooxml: 11 },
  tabloid: { label: 'Tabloid', w: 11, h: 17, ooxml: 3 }
}

export interface PageMargins {
  top: number
  bottom: number
  left: number
  right: number
  header: number
  footer: number
}

/** Excel's three built-in margin presets (inches), shown on Page Layout > Margins. */
export const MARGIN_PRESETS: Record<'normal' | 'wide' | 'narrow', PageMargins> = {
  normal: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
  wide: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 },
  narrow: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 }
}

export interface HeaderFooterText {
  left: string
  center: string
  right: string
}

export function emptyHFText(): HeaderFooterText {
  return { left: '', center: '', right: '' }
}

export interface HeaderFooterSetup {
  oddHeader: HeaderFooterText
  oddFooter: HeaderFooterText
  evenHeader: HeaderFooterText
  evenFooter: HeaderFooterText
  firstHeader: HeaderFooterText
  firstFooter: HeaderFooterText
  differentOddEven: boolean
  differentFirst: boolean
  /** "Scale with document" / "Align with page margins", the two checkboxes at the bottom of the
   *  Header/Footer Page Setup tab. */
  scaleWithDoc: boolean
  alignWithMargins: boolean
}

export function defaultHeaderFooter(): HeaderFooterSetup {
  return {
    oddHeader: emptyHFText(), oddFooter: emptyHFText(),
    evenHeader: emptyHFText(), evenFooter: emptyHFText(),
    firstHeader: emptyHFText(), firstFooter: emptyHFText(),
    differentOddEven: false, differentFirst: false,
    scaleWithDoc: true, alignWithMargins: true
  }
}

export interface PageSetup {
  orientation: 'portrait' | 'landscape'
  paperSize: PaperSize
  /** true = fitToWidth/fitToHeight drive scaling ("Fit to"); false = scalePercent drives it ("Adjust to"). */
  fitToPage: boolean
  scalePercent: number // 10..400, used when !fitToPage
  fitToWidth: number // pages wide, 0 = not constrained ("Automatic")
  fitToHeight: number // pages tall, 0 = not constrained
  margins: PageMargins
  centerHorizontally: boolean
  centerVertically: boolean
  /** _xlnm.Print_Area: one or more ranges (Set/Add to Print Area can produce several disjoint blocks). */
  printArea: Range[]
  /** _xlnm.Print_Titles: rows/cols repeated on every printed page. */
  printTitleRows?: { r1: number; r2: number }
  printTitleCols?: { c1: number; c2: number }
  gridlines: boolean // print gridlines (Sheet Options > Gridlines > Print)
  headings: boolean // print row/column headings (Sheet Options > Headings > Print)
  blackAndWhite: boolean
  draft: boolean
  pageOrder: 'downThenOver' | 'overThenDown'
  useFirstPageNumber: boolean
  firstPageNumber: number
  headerFooter: HeaderFooterSetup
  /** Manual page breaks: row index = the row that starts a new page (break is above this row). */
  rowBreaks: number[]
  colBreaks: number[]
}

export function defaultPageSetup(): PageSetup {
  return {
    orientation: 'portrait',
    paperSize: 'letter',
    fitToPage: false,
    scalePercent: 100,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { ...MARGIN_PRESETS.normal },
    centerHorizontally: false,
    centerVertically: false,
    printArea: [],
    gridlines: false,
    headings: false,
    blackAndWhite: false,
    draft: false,
    pageOrder: 'downThenOver',
    useFirstPageNumber: false,
    firstPageNumber: 1,
    headerFooter: defaultHeaderFooter(),
    rowBreaks: [],
    colBreaks: []
  }
}

/** Excel's built-in header/footer preset list (Page Setup > Header/Footer dropdowns), the same
 *  entries in the same order as Excel 2021, expressed as the section that gets filled (Excel
 *  always puts these in the center section except "Page 1 of ?" and the sheet-name-only one). */
export const HF_PRESETS: { label: string; center: string }[] = [
  { label: '(none)', center: '' },
  { label: 'Page 1', center: '&P' },
  { label: 'Page 1 of ?', center: '&P of &N' },
  { label: 'sheet name', center: '&A' },
  { label: 'file name', center: '&F' },
  { label: 'file name, page 1', center: '&F, &P' },
  { label: 'Page 1, date', center: '&P, &D' },
  { label: 'file name, tab, page 1, date', center: '&F, &A, &P, &D' },
  { label: 'author, page 1, file name', center: 'author, &P, &F' }
]

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

/** Anchors the chart's top-left corner to a cell, offset by pixels within that cell (at 100% zoom). */
export interface ChartAnchor {
  r: number
  c: number
  dx: number
  dy: number
}

export interface ChartSpec {
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
}

// ---------- Drawing objects: pictures, shapes, text boxes ----------
export type DrawingObjectKind = 'picture' | 'shape' | 'textbox' | 'preserved'

/** Shapes offered by the Insert > Shapes gallery (a small but representative subset of Excel's,
 *  covering Lines, Rectangles, Basic Shapes, Block Arrows and Flowchart). Maps 1:1 to a DrawingML
 *  `prstGeom prst`, except 'line' which writes `<a:cxnSp>`-free `<a:prstGeom prst="line">` on an
 *  `xdr:sp` with zero-height as Excel itself does for a plain line shape. */
export type ShapeKind =
  | 'line' | 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'diamond' | 'pentagon' | 'hexagon' | 'octagon'
  | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'leftRightArrow' | 'chevron'
  | 'flowChartProcess' | 'flowChartDecision' | 'flowChartTerminator' | 'flowChartDocument' | 'flowChartConnector'
  | 'star5' | 'heart' | 'cloud' | 'smileyFace'

export interface DrawingObject {
  id: string
  kind: DrawingObjectKind
  name: string // shown in the Selection Pane / Name Box, e.g. "Picture 1", "Rectangle 2", "Text Box 3"
  anchor: ChartAnchor
  width: number // px at 100% zoom
  height: number
  rotation?: number // degrees, clockwise
  altText?: string

  // kind === 'picture'
  imageDataUrl?: string // data: URL (base64) - the embedded xl/media part's bytes
  /** Basic rectangular crop, each side a 0..100 percent of the original image trimmed away
   *  (DrawingML `<a:srcRect>`, in thousandths of a percent). */
  cropPct?: { l: number; t: number; r: number; b: number }

  // kind === 'shape' | 'textbox'
  shapeKind?: ShapeKind // only for 'shape'; 'textbox' is always an `xdr:sp` with txBox="1"
  text?: string
  fillColor?: string | null // hex '#RRGGBB'; null = no fill
  lineColor?: string | null // hex '#RRGGBB'; null = no line
  lineWidth?: number // px
  textColor?: string // hex '#RRGGBB'

  // kind === 'preserved': an anchor we read but can't model (unsupported chart type, SmartArt,
  // form control, grouped shape, OLE object, ...) - kept verbatim so save never loses it.
  /** The raw `<xdr:...Anchor>...</xdr:...Anchor>` element exactly as read, except every `r:id`
   *  attribute value is replaced with the token `{{REL:<n>}}` (n indexing preservedRels) so fresh
   *  relationship ids can be assigned on write without colliding with this sheet's other objects. */
  preservedAnchorXml?: string
  /** The relationships this preserved anchor's `{{REL:n}}` tokens need, and (one hop further) the
   *  parts those referenced parts themselves reference (e.g. a preserved chart's own colors/style
   *  parts) - everything needed to round-trip the object's referenced parts unchanged. */
  preservedRels?: PreservedRel[]
}

/** One relationship a preserved drawing object depends on, plus the target part's own bytes so it
 *  can be re-added to the package on save (its original zip path is kept only for content-type
 *  lookup; the actual path it's written back at is chosen fresh at save time to avoid colliding
 *  with newly-created chart/media parts - see xlsxCharts.ts). */
export interface PreservedRel {
  /** Index referenced by `{{REL:n}}` tokens in preservedAnchorXml or in another PreservedRel's own part. */
  n: number
  relType: string // the Relationship's Type attribute
  originalPath: string // original zip path, e.g. 'xl/charts/chart4.xml' or 'xl/media/image3.png'
  contentType?: string // from [Content_Types].xml, when it's an Override (not a Default extension)
  dataBase64: string // the part's raw bytes, base64-encoded
  /** This part's own relationships (its `_rels/<part>.rels`), one hop only - e.g. a chart part's
   *  rels to its colors/style parts or embedded media. Each nested rel's `n` is also referenced by
   *  `{{REL:n}}` tokens inside this rel's own dataBase64 (decoded as text) when it's XML. */
  nestedRels?: PreservedRel[]
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

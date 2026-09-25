// DrawingML chart parts (xl/drawings/drawingN.xml, xl/charts/chartN.xml) for ChartSpec (model/types.ts).
// Write side is hooked into the raw-package post-processing step (xlsxPostProcess.ts) the same way
// AutoFilter/VBA are: exceljs has no concept of charts at all, so the parts are built here as plain
// strings and injected into the zip after exceljs writes the rest of the package.
// Read side (parseChartsForSheet) is called from the fast reader (fastXlsx.ts) and must stay free of
// exceljs/hyperformula/react imports, same constraint as the rest of that reader.
import type { CellValue, ChartAnchor, ChartElementsVisible, ChartSpec, ChartSubtype, ChartType, LegendPosition, Range, SheetLayout } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, colName, isError, normRange } from '../model/types'
import { colorSet } from '../charts/theme'
import { colLetterToIndex } from './xlsxShared'
import { decodeXmlEntities, parseRelationships, resolveRelTarget } from './fastXlsxXml'

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const DRAWING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'

// Axis ids: fixed, arbitrary but unique per chart (each chart part is self-contained, so reusing
// the same ids across charts is fine).
const CAT_AXIS_ID = 111111111
const VAL_AXIS_ID = 222222222
const X_AXIS_ID = 333333333
const Y_AXIS_ID = 444444444

// ---------- shared helpers ----------
export function pxToEmu(px: number): number {
  return Math.round(px * 9525)
}
export function emuToPx(emu: number): number {
  return emu / 9525
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function quoteSheetName(name: string): string {
  const safe = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)
  return safe ? name : `'${name.replace(/'/g, "''")}'`
}

function cellRefAbs(r: number, c: number): string {
  return '$' + colName(c) + '$' + (r + 1)
}
function rangeRefAbs(rg: Range): string {
  const n = normRange(rg)
  return n.r1 === n.r2 && n.c1 === n.c2 ? cellRefAbs(n.r1, n.c1) : cellRefAbs(n.r1, n.c1) + ':' + cellRefAbs(n.r2, n.c2)
}
function formulaRef(sheetName: string, rg: Range): string {
  return quoteSheetName(sheetName) + '!' + rangeRefAbs(rg)
}

interface SizeMaps {
  colWidths: Record<number, number>
  rowHeights: Record<number, number>
}
function colWidthPx(maps: SizeMaps, c: number): number {
  return maps.colWidths[c] ?? DEFAULT_COL_WIDTH
}
function rowHeightPx(maps: SizeMaps, r: number): number {
  return maps.rowHeights[r] ?? DEFAULT_ROW_HEIGHT
}

/** Walks forward `sizePx` from (startIdx, startOffsetPx), returning the resulting cell + offset -
 *  used to compute a twoCellAnchor's `to` cell from the chart's stored pixel width/height. */
function walkForward(startIdx: number, startOffsetPx: number, sizePx: number, sizeOf: (i: number) => number): { idx: number; offsetPx: number } {
  let idx = startIdx
  let remaining = startOffsetPx + sizePx
  for (let guard = 0; guard < 1_000_000; guard++) {
    const cell = Math.max(1, sizeOf(idx))
    if (remaining < cell) return { idx, offsetPx: remaining }
    remaining -= cell
    idx++
  }
  return { idx, offsetPx: remaining }
}

/** Inverse of walkForward: the pixel distance between two (cell, offsetPx) positions - used to
 *  recover a twoCellAnchor's width/height in px on read. */
function distancePx(idx1: number, off1: number, idx2: number, off2: number, sizeOf: (i: number) => number): number {
  if (idx2 <= idx1) return Math.max(0, off2 - off1)
  let total = sizeOf(idx1) - off1
  for (let i = idx1 + 1; i < idx2; i++) total += sizeOf(i)
  total += off2
  return Math.max(0, total)
}

// ---------- reading a grid of literal cell values (write side cached values) ----------
// Mirrors charts/dataExtraction.ts's own numOf/textOf/isBlank/header-detection logic exactly (kept
// as a small local copy rather than importing it, since that module reads live CellValue from the
// engine and we only have each cell's raw literal text at write time - a formula cell's raw text
// starts with '=' and reads as blank/unknown here, same as any other reader with no engine).
function cellValueFromRaw(raw: string | undefined): CellValue {
  if (raw === undefined || raw === '') return null
  if (raw.charCodeAt(0) === 61 /* '=' */) return null
  if (raw === 'TRUE') return true
  if (raw === 'FALSE') return false
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw)) return Number(raw)
  return raw
}
function numOf(v: CellValue): number | null {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  return null
}
function textOf(v: CellValue): string {
  if (v === null || v === undefined) return ''
  if (isError(v)) return v.error
  return String(v)
}
function isBlank(v: CellValue): boolean {
  return v === null || v === undefined || v === ''
}
function rowIsAllText(get: (r: number, c: number) => CellValue, r: number, c1: number, c2: number): boolean {
  let any = false
  for (let c = c1; c <= c2; c++) {
    const v = get(r, c)
    if (isBlank(v)) continue
    if (typeof v !== 'string') return false
    any = true
  }
  return any
}
function colIsAllText(get: (r: number, c: number) => CellValue, c: number, r1: number, r2: number): boolean {
  let any = false
  for (let r = r1; r <= r2; r++) {
    const v = get(r, c)
    if (isBlank(v)) continue
    if (typeof v !== 'string') return false
    any = true
  }
  return any
}

interface AxisSeries {
  name: string
  nameRef?: Range
  valRange: Range
  values: (number | null)[]
}
interface AxisData {
  categories: string[]
  categoryRange?: Range
  series: AxisSeries[]
}

/** Same layout (header detection, series/category split) as charts/dataExtraction.ts's
 *  extractChartData, but also returns the cell ranges backing each series/category so the chart
 *  XML's formulas can point at them, not just their extracted text/numbers. */
function buildAxisData(dataRange: Range, seriesInColumns: boolean, cells: Record<string, { raw?: string }>): AxisData {
  const get = (r: number, c: number): CellValue => cellValueFromRaw(cells[r + ',' + c]?.raw)
  const { r1, c1, r2, c2 } = normRange(dataRange)
  const rows = r2 - r1 + 1
  const cols = c2 - c1 + 1

  const hasColHeader = rows > 1 && rowIsAllText(get, r1, c1 + 1, c2)
  const hasRowHeader = cols > 1 && colIsAllText(get, c1, r1 + 1, r2)

  const dataR1 = r1 + (hasColHeader ? 1 : 0)
  const dataC1 = c1 + (hasRowHeader ? 1 : 0)

  const categories: string[] = []
  let categoryRange: Range | undefined
  const series: AxisSeries[] = []

  if (seriesInColumns) {
    for (let r = dataR1; r <= r2; r++) categories.push(hasRowHeader ? textOf(get(r, c1)) : String(r - r1 + 1))
    if (hasRowHeader) categoryRange = { r1: dataR1, c1, r2, c2: c1 }
    for (let c = dataC1; c <= c2; c++) {
      const name = hasColHeader ? textOf(get(r1, c)) : `Series${c - dataC1 + 1}`
      const values: (number | null)[] = []
      for (let r = dataR1; r <= r2; r++) values.push(numOf(get(r, c)))
      series.push({ name, nameRef: hasColHeader ? { r1, c1: c, r2: r1, c2: c } : undefined, valRange: { r1: dataR1, c1: c, r2, c2: c }, values })
    }
  } else {
    for (let c = dataC1; c <= c2; c++) categories.push(hasColHeader ? textOf(get(r1, c)) : String(c - c1 + 1))
    if (hasColHeader) categoryRange = { r1, c1: dataC1, r2: r1, c2 }
    for (let r = dataR1; r <= r2; r++) {
      const name = hasRowHeader ? textOf(get(r, c1)) : `Series${r - dataR1 + 1}`
      const values: (number | null)[] = []
      for (let c = dataC1; c <= c2; c++) values.push(numOf(get(r, c)))
      series.push({ name, nameRef: hasRowHeader ? { r1: r, c1, r2: r, c2: c1 } : undefined, valRange: { r1: r, c1: dataC1, r2: r, c2 }, values })
    }
  }
  return { categories, categoryRange, series }
}

// ---------- chart type <-> subtype mapping ----------
type ChartFamily = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter'

function chartFamily(subtype: ChartSubtype): ChartFamily {
  switch (subtype) {
    case 'clusteredColumn': case 'stackedColumn': case 'percentStackedColumn':
    case 'clusteredBar': case 'stackedBar': case 'percentStackedBar':
      return 'bar'
    case 'line': case 'stackedLine': case 'lineMarkers':
      return 'line'
    case 'area': case 'stackedArea':
      return 'area'
    case 'pie':
      return 'pie'
    case 'doughnut':
      return 'doughnut'
    case 'scatter': case 'scatterSmooth':
      return 'scatter'
  }
}

function barDirGrouping(subtype: ChartSubtype): { dir: 'col' | 'bar'; grouping: 'clustered' | 'stacked' | 'percentStacked' } {
  switch (subtype) {
    case 'stackedColumn': return { dir: 'col', grouping: 'stacked' }
    case 'percentStackedColumn': return { dir: 'col', grouping: 'percentStacked' }
    case 'clusteredBar': return { dir: 'bar', grouping: 'clustered' }
    case 'stackedBar': return { dir: 'bar', grouping: 'stacked' }
    case 'percentStackedBar': return { dir: 'bar', grouping: 'percentStacked' }
    default: return { dir: 'col', grouping: 'clustered' }
  }
}
function lineGroupingMarker(subtype: ChartSubtype): { grouping: 'standard' | 'stacked'; marker: boolean } {
  switch (subtype) {
    case 'stackedLine': return { grouping: 'stacked', marker: false }
    case 'lineMarkers': return { grouping: 'standard', marker: true }
    default: return { grouping: 'standard', marker: false }
  }
}
function areaGrouping(subtype: ChartSubtype): 'standard' | 'stacked' {
  return subtype === 'stackedArea' ? 'stacked' : 'standard'
}

// ---------- chart XML fragment builders ----------
function numRefOrLitXml(tag: string, sheetName: string, range: Range | undefined, values: (number | null)[]): string {
  const pts = values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')
  const body = `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}`
  if (range) {
    return `<c:${tag}><c:numRef><c:f>${escapeXml(formulaRef(sheetName, range))}</c:f><c:numCache>${body}</c:numCache></c:numRef></c:${tag}>`
  }
  return `<c:${tag}><c:numLit>${body}</c:numLit></c:${tag}>`
}

function catXml(sheetName: string, categoryRange: Range | undefined, categories: string[]): string {
  if (categoryRange) {
    const pts = categories.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join('')
    return `<c:cat><c:strRef><c:f>${escapeXml(formulaRef(sheetName, categoryRange))}</c:f><c:strCache><c:ptCount val="${categories.length}"/>${pts}</c:strCache></c:strRef></c:cat>`
  }
  const pts = categories.map((_, i) => `<c:pt idx="${i}"><c:v>${i + 1}</c:v></c:pt>`).join('')
  return `<c:cat><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${categories.length}"/>${pts}</c:numLit></c:cat>`
}

function seriesCommonXml(idx: number, name: string, nameRef: Range | undefined, sheetName: string, colorHex: string, useLine = false): string {
  const hex = colorHex.replace(/^#/, '')
  const tx = nameRef
    ? `<c:tx><c:strRef><c:f>${escapeXml(formulaRef(sheetName, nameRef))}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
    : ''
  const spPr = useLine
    ? `<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr>`
    : `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>`
  return `<c:idx val="${idx}"/><c:order val="${idx}"/>${tx}${spPr}`
}

function dLblsXml(): string {
  return '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
}

function richTextXml(text: string): string {
  return `<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(text)}</a:t></a:r></a:p>`
}
function titleXml(text: string): string {
  return `<c:title><c:tx><c:rich>${richTextXml(text)}</c:rich></c:tx><c:overlay val="0"/></c:title>`
}
function legendXml(pos: LegendPosition): string {
  const code = pos === 'right' ? 'r' : pos === 'top' ? 't' : pos === 'bottom' ? 'b' : pos === 'left' ? 'l' : undefined
  if (!code) return ''
  return `<c:legend><c:legendPos val="${code}"/><c:overlay val="0"/></c:legend>`
}

function catAxXml(chart: ChartSpec): string {
  const deleted = chart.elements.axes ? 0 : 1
  const title = chart.elements.axisTitles && chart.axisTitles?.x ? titleXml(chart.axisTitles.x) : ''
  return `<c:catAx><c:axId val="${CAT_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted}"/><c:axPos val="b"/>${title}<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${VAL_AXIS_ID}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
}
function valAxXml(chart: ChartSpec): string {
  const deleted = chart.elements.axes ? 0 : 1
  const title = chart.elements.axisTitles && chart.axisTitles?.y ? titleXml(chart.axisTitles.y) : ''
  const grid = chart.elements.gridlines ? '<c:majorGridlines/>' : ''
  return `<c:valAx><c:axId val="${VAL_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted}"/><c:axPos val="l"/>${grid}${title}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${CAT_AXIS_ID}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
}
function scatterAxXml(chart: ChartSpec): string {
  const deleted = chart.elements.axes ? 0 : 1
  const titleX = chart.elements.axisTitles && chart.axisTitles?.x ? titleXml(chart.axisTitles.x) : ''
  const titleY = chart.elements.axisTitles && chart.axisTitles?.y ? titleXml(chart.axisTitles.y) : ''
  const grid = chart.elements.gridlines ? '<c:majorGridlines/>' : ''
  const x = `<c:valAx><c:axId val="${X_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted}"/><c:axPos val="b"/>${titleX}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${Y_AXIS_ID}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
  const y = `<c:valAx><c:axId val="${Y_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted}"/><c:axPos val="l"/>${grid}${titleY}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${X_AXIS_ID}"/><c:crosses val="autoZero"/></c:valAx>`
  return x + y
}

/** Builds one xl/charts/chartN.xml document (a full c:chartSpace) for `chart`, reading cached
 *  series values from `cells` (the sheet's literal cell text - see cellValueFromRaw). */
export function buildChartXml(chart: ChartSpec, sheetName: string, cells: Record<string, { raw?: string }>): string {
  const family = chartFamily(chart.subtype)
  const axisData = buildAxisData(chart.dataRange, chart.seriesInColumns, cells)
  const colors = colorSet(chart.colorSetId ?? 0, Math.max(1, axisData.series.length))
  const colorAt = (i: number): string => colors[i % colors.length]

  let groupInner = ''
  let axesXml = ''

  if (family === 'bar') {
    const { dir, grouping } = barDirGrouping(chart.subtype)
    const overlap = grouping === 'clustered' ? '' : '<c:overlap val="100"/>'
    const seriesXml = axisData.series
      .map((s, i) => {
        const common = seriesCommonXml(i, s.name, s.nameRef, sheetName, colorAt(i))
        return `<c:ser>${common}${catXml(sheetName, axisData.categoryRange, axisData.categories)}${numRefOrLitXml('val', sheetName, s.valRange, s.values)}</c:ser>`
      })
      .join('')
    groupInner = `<c:barChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${chart.elements.dataLabels ? dLblsXml() : ''}<c:gapWidth val="150"/>${overlap}<c:axId val="${CAT_AXIS_ID}"/><c:axId val="${VAL_AXIS_ID}"/></c:barChart>`
    axesXml = catAxXml(chart) + valAxXml(chart)
  } else if (family === 'line') {
    const { grouping, marker } = lineGroupingMarker(chart.subtype)
    const seriesXml = axisData.series
      .map((s, i) => {
        const common = seriesCommonXml(i, s.name, s.nameRef, sheetName, colorAt(i), true)
        const markerXml = marker
          ? `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${colorAt(i).replace('#', '')}"/></a:solidFill></c:spPr></c:marker>`
          : '<c:marker><c:symbol val="none"/></c:marker>'
        return `<c:ser>${common}${markerXml}${catXml(sheetName, axisData.categoryRange, axisData.categories)}${numRefOrLitXml('val', sheetName, s.valRange, s.values)}<c:smooth val="0"/></c:ser>`
      })
      .join('')
    groupInner = `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${chart.elements.dataLabels ? dLblsXml() : ''}<c:marker val="1"/><c:axId val="${CAT_AXIS_ID}"/><c:axId val="${VAL_AXIS_ID}"/></c:lineChart>`
    axesXml = catAxXml(chart) + valAxXml(chart)
  } else if (family === 'area') {
    const grouping = areaGrouping(chart.subtype)
    const seriesXml = axisData.series
      .map((s, i) => {
        const common = seriesCommonXml(i, s.name, s.nameRef, sheetName, colorAt(i))
        return `<c:ser>${common}${catXml(sheetName, axisData.categoryRange, axisData.categories)}${numRefOrLitXml('val', sheetName, s.valRange, s.values)}</c:ser>`
      })
      .join('')
    groupInner = `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${chart.elements.dataLabels ? dLblsXml() : ''}<c:axId val="${CAT_AXIS_ID}"/><c:axId val="${VAL_AXIS_ID}"/></c:areaChart>`
    axesXml = catAxXml(chart) + valAxXml(chart)
  } else if (family === 'pie' || family === 'doughnut') {
    const s = axisData.series[0]
    const seriesXml = s
      ? (() => {
          const common = seriesCommonXml(0, s.name, s.nameRef, sheetName, colorAt(0))
          const dPts = axisData.categories
            .map((_, idx) => `<c:dPt><c:idx val="${idx}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${colorAt(idx).replace('#', '')}"/></a:solidFill></c:spPr></c:dPt>`)
            .join('')
          const dLbls = chart.elements.dataLabels ? dLblsXml() : ''
          return `<c:ser>${common}${dPts}${dLbls}${catXml(sheetName, axisData.categoryRange, axisData.categories)}${numRefOrLitXml('val', sheetName, s.valRange, s.values)}</c:ser>`
        })()
      : ''
    const tail = family === 'doughnut' ? '<c:firstSliceAng val="0"/><c:holeSize val="50"/>' : '<c:firstSliceAng val="0"/>'
    const tag = family === 'pie' ? 'pieChart' : 'doughnutChart'
    groupInner = `<c:${tag}><c:varyColors val="1"/>${seriesXml}${tail}</c:${tag}>`
  } else {
    // scatter
    const scatterStyle = chart.subtype === 'scatterSmooth' ? 'smoothMarker' : 'marker'
    const smoothFlag = chart.subtype === 'scatterSmooth'
    const seriesXml = axisData.series
      .map((s, i) => {
        const hex = colorAt(i).replace('#', '')
        const common = seriesCommonXml(i, s.name, s.nameRef, sheetName, hex, smoothFlag)
        const markerXml = `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr></c:marker>`
        const xValues = axisData.categories.map((v, idx) => {
          const n = parseFloat(v)
          return Number.isFinite(n) ? n : idx + 1
        })
        const xVal = numRefOrLitXml('xVal', sheetName, axisData.categoryRange, xValues)
        const yVal = numRefOrLitXml('yVal', sheetName, s.valRange, s.values)
        const smooth = smoothFlag ? '<c:smooth val="1"/>' : '<c:smooth val="0"/>'
        return `<c:ser>${common}${markerXml}${xVal}${yVal}${smooth}</c:ser>`
      })
      .join('')
    groupInner = `<c:scatterChart><c:scatterStyle val="${scatterStyle}"/><c:varyColors val="0"/>${seriesXml}${chart.elements.dataLabels ? dLblsXml() : ''}<c:axId val="${X_AXIS_ID}"/><c:axId val="${Y_AXIS_ID}"/></c:scatterChart>`
    axesXml = scatterAxXml(chart)
  }

  const showTitle = chart.elements.title && chart.title
  const titleEl = showTitle ? titleXml(chart.title) : ''
  const autoTitleDeleted = showTitle ? '' : '<c:autoTitleDeleted val="1"/>'
  const legendEl = chart.elements.legend ? legendXml(chart.legendPosition) : ''

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<c:chart>${titleEl}${autoTitleDeleted}<c:plotArea><c:layout/>${groupInner}${axesXml}</c:plotArea>${legendEl}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    '</c:chartSpace>'
  )
}

// ---------- drawing (anchor) XML ----------
function buildAnchorXml(chart: ChartSpec, maps: SizeMaps, rId: string, shapeId: number): string {
  const fromCol = chart.anchor.c
  const fromColOff = pxToEmu(chart.anchor.dx)
  const fromRow = chart.anchor.r
  const fromRowOff = pxToEmu(chart.anchor.dy)
  const toCol = walkForward(chart.anchor.c, chart.anchor.dx, chart.width, (c) => colWidthPx(maps, c))
  const toRow = walkForward(chart.anchor.r, chart.anchor.dy, chart.height, (r) => rowHeightPx(maps, r))
  return (
    '<xdr:twoCellAnchor>' +
    `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>${fromColOff}</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>${fromRowOff}</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${toCol.idx}</xdr:col><xdr:colOff>${pxToEmu(toCol.offsetPx)}</xdr:colOff><xdr:row>${toRow.idx}</xdr:row><xdr:rowOff>${pxToEmu(toRow.offsetPx)}</xdr:rowOff></xdr:to>` +
    '<xdr:graphicFrame macro="">' +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="Chart ${shapeId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${rId}"/></a:graphicData></a:graphic>` +
    '</xdr:graphicFrame>' +
    '<xdr:clientData/>' +
    '</xdr:twoCellAnchor>'
  )
}

function buildDrawingXml(anchorsXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}">${anchorsXml}</xdr:wsDr>`
}
function buildDrawingRelsXml(rels: { id: string; target: string }[]): string {
  const body = rels.map((r) => `<Relationship Id="${r.id}" Type="${CHART_REL_TYPE}" Target="${r.target}"/>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`
}

// ---------- write: inject drawings/charts for every sheet that has them ----------
/** Patches `files` (the unzipped .xlsx package exceljs just wrote) in place: for each sheet with
 *  `layout.charts`, adds xl/drawings/drawingN.xml (+ its rels), one xl/charts/chartN.xml per
 *  chart, a `<drawing r:id>` in the sheet XML (schema position: right before `<legacyDrawing>` if
 *  the sheet has cell notes/comments, else right before `</worksheet>`), a relationship in the
 *  sheet's own rels, and Content_Types overrides for the new parts. */
export function injectCharts(
  files: Record<string, Uint8Array>,
  sheets: { name: string; cells: Record<string, { raw?: string }>; layout?: Partial<SheetLayout> }[]
): void {
  const enc = new TextEncoder()
  const dec = new TextDecoder('utf-8')
  let drawingCounter = 0
  let chartCounter = 0
  const overridesToAdd: string[] = []

  sheets.forEach((sd, i) => {
    const charts = sd.layout?.charts
    if (!charts || !charts.length) return
    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
    const sheetXmlBytes = files[sheetPath]
    if (!sheetXmlBytes) return
    let xml = dec.decode(sheetXmlBytes)

    drawingCounter++
    const drawingName = `drawing${drawingCounter}.xml`
    const drawingPath = `xl/drawings/${drawingName}`
    const drawingRelsPath = `xl/drawings/_rels/${drawingName}.rels`

    const maps: SizeMaps = { colWidths: sd.layout?.colWidths ?? {}, rowHeights: sd.layout?.rowHeights ?? {} }
    const anchors: string[] = []
    const drawingRels: { id: string; target: string }[] = []

    charts.forEach((chart, ci) => {
      chartCounter++
      const chartName = `chart${chartCounter}.xml`
      files[`xl/charts/${chartName}`] = enc.encode(buildChartXml(chart, sd.name, sd.cells))
      const rId = `rId${ci + 1}`
      drawingRels.push({ id: rId, target: `../charts/${chartName}` })
      anchors.push(buildAnchorXml(chart, maps, rId, 1000 + ci))
      overridesToAdd.push(`<Override PartName="/xl/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`)
    })

    files[drawingPath] = enc.encode(buildDrawingXml(anchors.join('')))
    files[drawingRelsPath] = enc.encode(buildDrawingRelsXml(drawingRels))
    overridesToAdd.push(`<Override PartName="/xl/drawings/${drawingName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`)

    // Sheet's own rels: add (or create) a relationship to the drawing part.
    const relsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let relsXml = files[relsPath]
      ? dec.decode(files[relsPath])
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    let maxId = 0
    for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, parseInt(m[1], 10))
    const drawingRelId = `rId${maxId + 1}`
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${drawingRelId}" Type="${DRAWING_REL_TYPE}" Target="../drawings/${drawingName}"/></Relationships>`)
    files[relsPath] = enc.encode(relsXml)

    // Sheet XML: `<drawing>` sits right before `<legacyDrawing>` (comments/notes) in CT_Worksheet's
    // schema order, else right before `</worksheet>`.
    const drawingEl = `<drawing r:id="${drawingRelId}"/>`
    xml = xml.indexOf('<legacyDrawing') !== -1 ? xml.replace('<legacyDrawing', drawingEl + '<legacyDrawing') : xml.replace('</worksheet>', drawingEl + '</worksheet>')
    files[sheetPath] = enc.encode(xml)
  })

  if (overridesToAdd.length) {
    const ctPath = '[Content_Types].xml'
    const ctXml = dec.decode(files[ctPath])
    files[ctPath] = enc.encode(ctXml.replace('</Types>', overridesToAdd.join('') + '</Types>'))
  }
}

// ---------- read: drawings/charts -> ChartSpec[] ----------
function parseCellPos(inner: string): { col: number; colOffPx: number; row: number; rowOffPx: number } | null {
  const col = /<xdr:col>(\d+)<\/xdr:col>/.exec(inner)
  const row = /<xdr:row>(\d+)<\/xdr:row>/.exec(inner)
  if (!col || !row) return null
  const colOff = /<xdr:colOff>(-?\d+)<\/xdr:colOff>/.exec(inner)
  const rowOff = /<xdr:rowOff>(-?\d+)<\/xdr:rowOff>/.exec(inner)
  return {
    col: parseInt(col[1], 10),
    colOffPx: emuToPx(colOff ? parseInt(colOff[1], 10) : 0),
    row: parseInt(row[1], 10),
    rowOffPx: emuToPx(rowOff ? parseInt(rowOff[1], 10) : 0)
  }
}

interface AnchorInfo {
  anchor: ChartAnchor
  width: number
  height: number
}

function parseTwoCellAnchor(body: string, maps: SizeMaps): AnchorInfo | null {
  const fromM = /<xdr:from>([\s\S]*?)<\/xdr:from>/.exec(body)
  const toM = /<xdr:to>([\s\S]*?)<\/xdr:to>/.exec(body)
  if (!fromM || !toM) return null
  const from = parseCellPos(fromM[1])
  const to = parseCellPos(toM[1])
  if (!from || !to) return null
  const widthPx = distancePx(from.col, from.colOffPx, to.col, to.colOffPx, (c) => colWidthPx(maps, c))
  const heightPx = distancePx(from.row, from.rowOffPx, to.row, to.rowOffPx, (r) => rowHeightPx(maps, r))
  return {
    anchor: { r: from.row, c: from.col, dx: Math.round(from.colOffPx), dy: Math.round(from.rowOffPx) },
    width: Math.round(widthPx),
    height: Math.round(heightPx)
  }
}
function parseOneCellAnchor(body: string): AnchorInfo | null {
  const fromM = /<xdr:from>([\s\S]*?)<\/xdr:from>/.exec(body)
  const extM = /<xdr:ext\s+cx="(-?\d+)"\s+cy="(-?\d+)"/.exec(body)
  if (!fromM || !extM) return null
  const from = parseCellPos(fromM[1])
  if (!from) return null
  return {
    anchor: { r: from.row, c: from.col, dx: Math.round(from.colOffPx), dy: Math.round(from.rowOffPx) },
    width: Math.round(emuToPx(parseInt(extM[1], 10))),
    height: Math.round(emuToPx(parseInt(extM[2], 10)))
  }
}

function extractFormulaRef(block: string, tag: string): string | undefined {
  const outer = new RegExp(`<c:${tag}>([\\s\\S]*?)<\\/c:${tag}>`).exec(block)
  if (!outer) return undefined
  const f = /<c:f>([\s\S]*?)<\/c:f>/.exec(outer[1])
  return f ? decodeXmlEntities(f[1]) : undefined
}
function extractSeriesTxRef(block: string): string | undefined {
  return extractFormulaRef(block, 'tx')
}

function parseFormulaRangeRef(ref: string): Range | null {
  const bang = ref.lastIndexOf('!')
  const rangePart = bang === -1 ? ref : ref.slice(bang + 1)
  const cleaned = rangePart.replace(/\$/g, '')
  const m = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(cleaned)
  if (!m) return null
  const c1 = colLetterToIndex(m[1].toUpperCase())
  const r1 = parseInt(m[2], 10) - 1
  const c2 = m[3] ? colLetterToIndex(m[3].toUpperCase()) : c1
  const r2 = m[4] ? parseInt(m[4], 10) - 1 : r1
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }
}
function unionRange(a: Range | null, b: Range | null): Range | null {
  if (!a) return b
  if (!b) return a
  return { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) }
}

/** Chart types the schema supports that we have no ChartSpec representation for (combo/3D/other
 *  chart groups). A chart part using only these is skipped entirely by parseChartXml (returns null). */
const UNSUPPORTED_FAMILY_TAGS = [
  'radarChart', 'bubbleChart', 'stockChart', 'surfaceChart', 'surface3DChart',
  'bar3DChart', 'line3DChart', 'area3DChart', 'pie3DChart', 'ofPieChart'
]
const FAMILY_TAGS = ['barChart', 'lineChart', 'areaChart', 'pieChart', 'doughnutChart', 'scatterChart'] as const

function parseChartXml(xml: string, anchorInfo: AnchorInfo): ChartSpec | null {
  let family: (typeof FAMILY_TAGS)[number] | undefined
  for (const tag of FAMILY_TAGS) {
    if (xml.indexOf('<c:' + tag) !== -1) {
      family = tag
      break
    }
  }
  if (!family) return null // combo/3D/radar/bubble/stock/surface/ofPie - no ChartSpec equivalent

  const groupM = new RegExp(`<c:${family}\\b[^>]*>([\\s\\S]*?)<\\/c:${family}>`).exec(xml)
  const groupBody = groupM ? groupM[1] : ''

  let type: ChartType
  let subtype: ChartSubtype
  if (family === 'barChart') {
    const dir = /<c:barDir val="([^"]+)"/.exec(groupBody)?.[1] ?? 'col'
    const grouping = /<c:grouping val="([^"]+)"/.exec(groupBody)?.[1] ?? 'clustered'
    type = 'bar'
    subtype =
      dir === 'bar'
        ? grouping === 'stacked' ? 'stackedBar' : grouping === 'percentStacked' ? 'percentStackedBar' : 'clusteredBar'
        : grouping === 'stacked' ? 'stackedColumn' : grouping === 'percentStacked' ? 'percentStackedColumn' : 'clusteredColumn'
  } else if (family === 'lineChart') {
    const grouping = /<c:grouping val="([^"]+)"/.exec(groupBody)?.[1] ?? 'standard'
    type = 'line'
    const hasMarker = /<c:marker>\s*<c:symbol val="(?!none)/.test(groupBody)
    subtype = grouping === 'stacked' || grouping === 'percentStacked' ? 'stackedLine' : hasMarker ? 'lineMarkers' : 'line'
  } else if (family === 'areaChart') {
    const grouping = /<c:grouping val="([^"]+)"/.exec(groupBody)?.[1] ?? 'standard'
    type = 'line'
    subtype = grouping === 'stacked' || grouping === 'percentStacked' ? 'stackedArea' : 'area'
  } else if (family === 'pieChart') {
    type = 'pie'
    subtype = 'pie'
  } else if (family === 'doughnutChart') {
    type = 'pie'
    subtype = 'doughnut'
  } else {
    const style = /<c:scatterStyle val="([^"]+)"/.exec(groupBody)?.[1] ?? 'marker'
    type = 'scatter'
    subtype = /smooth/i.test(style) ? 'scatterSmooth' : 'scatter'
  }

  const isScatter = family === 'scatterChart'
  const isPieLike = family === 'pieChart' || family === 'doughnutChart'
  const valTag = isScatter ? 'yVal' : 'val'
  const catTag = isScatter ? 'xVal' : 'cat'

  const serBlocks = groupBody.match(/<c:ser>[\s\S]*?<\/c:ser>/g) ?? []
  let dataRange: Range | null = null
  let firstValRange: Range | null = null
  serBlocks.forEach((block, i) => {
    const catRef = extractFormulaRef(block, catTag)
    const valRef = extractFormulaRef(block, valTag)
    const txRef = extractSeriesTxRef(block)
    for (const ref of [catRef, valRef, txRef]) {
      if (!ref) continue
      const rg = parseFormulaRangeRef(ref)
      if (!rg) continue
      dataRange = unionRange(dataRange, rg)
      if (i === 0 && ref === valRef) firstValRange = rg
    }
  })
  if (!dataRange) return null // no cell-backed data at all (pure literal chart) - nothing to anchor a re-read on

  let seriesInColumns = true
  if (firstValRange) {
    const fv: Range = firstValRange
    const rows = fv.r2 - fv.r1 + 1
    const cols = fv.c2 - fv.c1 + 1
    if (rows === 1 && cols > 1) seriesInColumns = false
  }

  const titleM = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml)
  const hasTitle = !!titleM
  const title = titleM ? [...titleM[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXmlEntities(m[1])).join('') : ''

  const legendM = /<c:legend>([\s\S]*?)<\/c:legend>/.exec(xml)
  const legendPosCode = legendM ? /<c:legendPos val="([^"]+)"/.exec(legendM[1])?.[1] : undefined
  const legendPosition: LegendPosition =
    legendPosCode === 'r' ? 'right' : legendPosCode === 't' ? 'top' : legendPosCode === 'b' ? 'bottom' : legendPosCode === 'l' ? 'left' : legendM ? 'right' : 'none'

  const gridlines = /<c:majorGridlines/.test(xml)
  const dataLabels = /<c:dLbls>[\s\S]*?<c:showVal val="1"/.test(xml)

  let axesVisible = false
  let axisTitleX: string | undefined
  let axisTitleY: string | undefined
  if (!isPieLike) {
    const axBlocks = [...xml.matchAll(/<c:(catAx|valAx)>([\s\S]*?)<\/c:\1>/g)]
    axesVisible = axBlocks.length === 0 || !axBlocks.every((m) => /<c:delete val="1"/.test(m[2]))
    for (const m of axBlocks) {
      const body = m[2]
      const posCode = /<c:axPos val="([^"]+)"/.exec(body)?.[1]
      const titleM2 = /<c:title>([\s\S]*?)<\/c:title>/.exec(body)
      if (!titleM2) continue
      const t = [...titleM2[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m2) => decodeXmlEntities(m2[1])).join('')
      if (posCode === 'b') axisTitleX = t
      else if (posCode === 'l') axisTitleY = t
    }
  }
  const hasAxisTitles = !!(axisTitleX || axisTitleY)

  const elements: ChartElementsVisible = {
    title: hasTitle,
    legend: !!legendM,
    gridlines,
    dataLabels,
    axes: axesVisible,
    axisTitles: hasAxisTitles
  }

  const spec: ChartSpec = {
    id: '',
    type,
    subtype,
    dataRange,
    seriesInColumns,
    title,
    anchor: anchorInfo.anchor,
    width: anchorInfo.width,
    height: anchorInfo.height,
    legendPosition,
    elements
  }
  if (hasAxisTitles) spec.axisTitles = { x: axisTitleX, y: axisTitleY }
  return spec
}

/** Reads every chart anchored on `sheetPath`'s drawing part (if any), mapped back to ChartSpec.
 *  `colWidths`/`rowHeights` are that sheet's already-parsed layout (px), used to convert a
 *  twoCellAnchor's end cell back into a pixel width/height. Charts using an unsupported chart
 *  group (radar, bubble, stock, surface, 3D, combo, ofPie - see UNSUPPORTED_FAMILY_TAGS) are
 *  silently skipped, same as a chart part with no recognizable series data. */
export function parseChartsForSheet(
  files: Record<string, Uint8Array>,
  dec: TextDecoder,
  sheetPath: string,
  colWidths: Record<number, number>,
  rowHeights: Record<number, number>
): ChartSpec[] {
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const file = sheetPath.split('/').pop() ?? ''
  const relsPath = `${dir}/_rels/${file}.rels`
  const relsBytes = files[relsPath]
  if (!relsBytes) return []
  const relsXml = dec.decode(relsBytes)
  const drawingRel = parseRelationships(relsXml).find((r) => r.type === DRAWING_REL_TYPE)
  if (!drawingRel) return []
  const drawingPath = resolveRelTarget(sheetPath, drawingRel.target)
  const drawingBytes = files[drawingPath]
  if (!drawingBytes) return []
  const drawingXml = dec.decode(drawingBytes)

  const drawingDir = drawingPath.split('/').slice(0, -1).join('/')
  const drawingFile = drawingPath.split('/').pop() ?? ''
  const drawingRelsPath = `${drawingDir}/_rels/${drawingFile}.rels`
  const drawingRelsBytes = files[drawingRelsPath]
  const drawingRelsMap = new Map<string, string>()
  if (drawingRelsBytes) {
    for (const r of parseRelationships(dec.decode(drawingRelsBytes))) drawingRelsMap.set(r.id, r.target)
  }

  const maps: SizeMaps = { colWidths, rowHeights }
  const charts: ChartSpec[] = []
  let idCounter = 0

  const anchorRe = /<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g
  let m: RegExpExecArray | null
  while ((m = anchorRe.exec(drawingXml))) {
    const kind = m[1]
    const body = m[2]
    const chartRefM = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(body)
    if (!chartRefM) continue // a non-chart drawing object (picture/shape) - not ours to read
    const rId = chartRefM[1]
    const target = drawingRelsMap.get(rId)
    if (!target) continue
    const chartPath = resolveRelTarget(drawingPath, target)
    const chartBytes = files[chartPath]
    if (!chartBytes) continue
    const chartXml = dec.decode(chartBytes)

    const anchorInfo = kind === 'twoCellAnchor' ? parseTwoCellAnchor(body, maps) : parseOneCellAnchor(body)
    if (!anchorInfo) continue

    const spec = parseChartXml(chartXml, anchorInfo)
    if (spec) {
      idCounter++
      spec.id = 'chart_r' + idCounter
      charts.push(spec)
    }
  }
  return charts
}

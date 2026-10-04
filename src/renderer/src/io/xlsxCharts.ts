import type { ChartSpec, ChartSubtype, ChartType, LegendPosition, Range, SheetLayout } from '../model/types'
import { normRange, colName } from '../model/types'
import { extractChartData, CellReader } from '../charts/dataExtraction'
import { decodeXmlEntities, parseRelationships, resolveRelTarget } from './fastXlsxXml'

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeXmlText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function quoteSheetName(name: string): string {
  if (/^[A-Za-z0-9_]+$/.test(name)) return name
  return `'${name.replace(/'/g, "''")}'`
}

function a1CellRef(sheet: string, r: number, c: number): string {
  return `${quoteSheetName(sheet)}!$${colName(c)}$${r + 1}`
}

function a1RangeRef(sheet: string, r1: number, c1: number, r2: number, c2: number): string {
  if (r1 === r2 && c1 === c2) return a1CellRef(sheet, r1, c1)
  return `${quoteSheetName(sheet)}!$${colName(c1)}$${r1 + 1}:$${colName(c2)}$${r2 + 1}`
}

function mapLegendPos(pos: LegendPosition): string {
  if (pos === 'top') return 't'
  if (pos === 'bottom') return 'b'
  if (pos === 'left') return 'l'
  return 'r'
}

/** Builds the drawingml chartSpace XML for a ChartSpec. */
export function buildChartSpaceXml(
  spec: ChartSpec,
  sheetName: string,
  cells: CellReader
): string {
  const norm = normRange(spec.dataRange)
  const extracted = extractChartData(cells, spec.dataRange, spec.seriesInColumns, spec.firstIsCategories)

  const rows = norm.r2 - norm.r1 + 1
  const cols = norm.c2 - norm.c1 + 1
  const hasColHeader = rows > 1 && extracted.series.length > 0 && extracted.series[0].name !== 'Series1'
  const hasRowHeader = cols > 1 && extracted.categories.length > 0 && extracted.categories[0] !== '1'

  const dataR1 = norm.r1 + (hasColHeader ? 1 : 0)
  const dataC1 = norm.c1 + (hasRowHeader ? 1 : 0)

  // Title XML
  let titleXml = ''
  if (spec.elements.title && spec.title) {
    titleXml = `
      <c:title>
        <c:tx>
          <c:rich>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr><a:defRPr/></a:pPr>
              <a:r>
                <a:rPr lang="en-US"/>
                <a:t>${escapeXmlText(spec.title)}</a:t>
              </a:r>
            </a:p>
          </c:rich>
        </c:tx>
        <c:layout/>
        <c:overlay val="0"/>
      </c:title>`
  } else {
    titleXml = `<c:autoTitleDeleted val="1"/>`
  }

  // Category reference formula
  let catFormula = ''
  if (extracted.categories.length > 0) {
    if (spec.seriesInColumns) {
      if (hasRowHeader) {
        catFormula = a1RangeRef(sheetName, dataR1, norm.c1, norm.r2, norm.c1)
      }
    } else {
      if (hasColHeader) {
        catFormula = a1RangeRef(sheetName, norm.r1, dataC1, norm.r1, norm.c2)
      }
    }
  }

  const isScatter = spec.type === 'scatter'
  const isPie = spec.subtype === 'pie'
  const isDoughnut = spec.subtype === 'doughnut'
  const isBar = spec.subtype.includes('Bar')
  const isLine = spec.type === 'line' && (spec.subtype === 'line' || spec.subtype === 'stackedLine' || spec.subtype === 'lineMarkers')
  const isArea = spec.subtype === 'area' || spec.subtype === 'stackedArea'
  const hasMarkers = spec.subtype === 'lineMarkers' || spec.subtype === 'scatter' || spec.subtype === 'scatterSmooth'
  const isStacked = spec.subtype.includes('stacked') || spec.subtype.includes('Stacked')
  const isPercent = spec.subtype.includes('percent') || spec.subtype.includes('Percent')

  // Series XML
  const seriesXmlList: string[] = []
  extracted.series.forEach((s, i) => {
    let nameFormula = ''
    let valFormula = ''

    if (spec.seriesInColumns) {
      const col = dataC1 + i
      if (hasColHeader) nameFormula = a1CellRef(sheetName, norm.r1, col)
      valFormula = a1RangeRef(sheetName, dataR1, col, norm.r2, col)
    } else {
      const row = dataR1 + i
      if (hasRowHeader) nameFormula = a1CellRef(sheetName, row, norm.c1)
      valFormula = a1RangeRef(sheetName, row, dataC1, row, norm.c2)
    }

    const txXml = nameFormula
      ? `<c:tx><c:strRef><c:f>${escapeXmlText(nameFormula)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXmlText(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
      : `<c:tx><c:v>${escapeXmlText(s.name)}</c:v></c:tx>`

    if (isScatter) {
      const xValXml = catFormula
        ? `<c:xVal><c:numRef><c:f>${escapeXmlText(catFormula)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${extracted.categories.length}"/>${extracted.categories.map((c, idx) => `<c:pt idx="${idx}"><c:v>${Number(c) || idx + 1}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:xVal>`
        : `<c:xVal><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${extracted.categories.length}"/>${extracted.categories.map((c, idx) => `<c:pt idx="${idx}"><c:v>${Number(c) || idx + 1}</c:v></c:pt>`).join('')}</c:numLit></c:xVal>`

      const yValXml = `<c:yVal><c:numRef><c:f>${escapeXmlText(valFormula)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${s.values.length}"/>${s.values.map((v, idx) => `<c:pt idx="${idx}"><c:v>${v !== null ? v : 0}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:yVal>`

      seriesXmlList.push(`
        <c:ser>
          <c:idx val="${i}"/>
          <c:order val="${i}"/>
          ${txXml}
          ${hasMarkers ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : ''}
          ${xValXml}
          ${yValXml}
        </c:ser>`)
    } else {
      const catXml = catFormula
        ? `<c:cat><c:strRef><c:f>${escapeXmlText(catFormula)}</c:f><c:strCache><c:ptCount val="${extracted.categories.length}"/>${extracted.categories.map((c, idx) => `<c:pt idx="${idx}"><c:v>${escapeXmlText(c)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`
        : extracted.categories.length > 0
          ? `<c:cat><c:strLit><c:ptCount val="${extracted.categories.length}"/>${extracted.categories.map((c, idx) => `<c:pt idx="${idx}"><c:v>${escapeXmlText(c)}</c:v></c:pt>`).join('')}</c:strLit></c:cat>`
          : ''

      const valXml = `<c:val><c:numRef><c:f>${escapeXmlText(valFormula)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${s.values.length}"/>${s.values.map((v, idx) => `<c:pt idx="${idx}"><c:v>${v !== null ? v : 0}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val>`

      seriesXmlList.push(`
        <c:ser>
          <c:idx val="${i}"/>
          <c:order val="${i}"/>
          ${txXml}
          ${hasMarkers ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : ''}
          ${catXml}
          ${valXml}
        </c:ser>`)
    }
  })

  // Plot Element
  let plotXml = ''
  const ax1 = 12345678
  const ax2 = 12345679

  if (spec.type === 'bar') {
    const grouping = isPercent ? 'percentStacked' : isStacked ? 'stacked' : 'clustered'
    const barDir = isBar ? 'bar' : 'col'
    plotXml = `
      <c:barChart>
        <c:barDir val="${barDir}"/>
        <c:grouping val="${grouping}"/>
        <c:varyColors val="0"/>
        ${seriesXmlList.join('')}
        ${isStacked || isPercent ? '<c:overlap val="100"/>' : ''}
        <c:axId val="${ax1}"/>
        <c:axId val="${ax2}"/>
      </c:barChart>`
  } else if (isLine) {
    const grouping = isStacked ? 'stacked' : 'standard'
    plotXml = `
      <c:lineChart>
        <c:grouping val="${grouping}"/>
        <c:varyColors val="0"/>
        ${seriesXmlList.join('')}
        <c:axId val="${ax1}"/>
        <c:axId val="${ax2}"/>
      </c:lineChart>`
  } else if (isArea) {
    const grouping = isStacked ? 'stacked' : 'standard'
    plotXml = `
      <c:areaChart>
        <c:grouping val="${grouping}"/>
        <c:varyColors val="0"/>
        ${seriesXmlList.join('')}
        <c:axId val="${ax1}"/>
        <c:axId val="${ax2}"/>
      </c:areaChart>`
  } else if (isPie) {
    plotXml = `
      <c:pieChart>
        <c:varyColors val="1"/>
        ${seriesXmlList.join('')}
      </c:pieChart>`
  } else if (isDoughnut) {
    plotXml = `
      <c:doughnutChart>
        <c:varyColors val="1"/>
        ${seriesXmlList.join('')}
        <c:holeSize val="75"/>
      </c:doughnutChart>`
  } else if (isScatter) {
    const style = spec.subtype === 'scatterSmooth' ? 'smoothMarker' : 'lineMarker'
    plotXml = `
      <c:scatterChart>
        <c:scatterStyle val="${style}"/>
        <c:varyColors val="0"/>
        ${seriesXmlList.join('')}
        <c:axId val="${ax1}"/>
        <c:axId val="${ax2}"/>
      </c:scatterChart>`
  }

  // Axes XML
  let axesXml = ''
  if (!isPie && !isDoughnut) {
    const gridlinesXml = spec.elements.gridlines ? '<c:majorGridlines/>' : ''
    if (isScatter) {
      axesXml = `
        <c:valAx>
          <c:axId val="${ax1}"/>
          <c:scaling><c:orientation val="minMax"/></c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          ${gridlinesXml}
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="${ax2}"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="midCat"/>
        </c:valAx>
        <c:valAx>
          <c:axId val="${ax2}"/>
          <c:scaling><c:orientation val="minMax"/></c:scaling>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          ${gridlinesXml}
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="${ax1}"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="midCat"/>
        </c:valAx>`
    } else {
      const catPos = isBar ? 'l' : 'b'
      const valPos = isBar ? 'b' : 'l'
      axesXml = `
        <c:catAx>
          <c:axId val="${ax1}"/>
          <c:scaling><c:orientation val="minMax"/></c:scaling>
          <c:delete val="0"/>
          <c:axPos val="${catPos}"/>
          ${gridlinesXml}
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="${ax2}"/>
          <c:crosses val="autoZero"/>
          <c:auto val="1"/>
          <c:lblAlgn val="ctr"/>
          <c:lblOffset val="100"/>
        </c:catAx>
        <c:valAx>
          <c:axId val="${ax2}"/>
          <c:scaling><c:orientation val="minMax"/></c:scaling>
          <c:delete val="0"/>
          <c:axPos val="${valPos}"/>
          ${gridlinesXml}
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="${ax1}"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="between"/>
        </c:valAx>`
    }
  }

  // Legend XML
  let legendXml = ''
  if (spec.elements.legend && spec.legendPosition !== 'none') {
    legendXml = `
      <c:legend>
        <c:legendPos val="${mapLegendPos(spec.legendPosition)}"/>
        <c:layout/>
        <c:overlay val="0"/>
      </c:legend>`
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    ${titleXml}
    <c:plotArea>
      <c:layout/>
      ${plotXml}
      ${axesXml}
    </c:plotArea>
    ${legendXml}
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
</c:chartSpace>`
}

/** Calculates the bottom-right cell and offsets for twoCellAnchor. */
export function computeToAnchor(
  anchor: ChartSpec['anchor'],
  width: number,
  height: number,
  layout?: Partial<SheetLayout>
): { toCol: number; toColOff: number; toRow: number; toRowOff: number } {
  let curC = anchor.c
  let remW = width + anchor.dx
  while (true) {
    const colW = layout?.colWidths?.[curC] ?? 64
    if (remW <= colW) break
    remW -= colW
    curC++
  }
  const toCol = curC
  const toColOff = Math.round(remW * 9525)

  let curR = anchor.r
  let remH = height + anchor.dy
  while (true) {
    const rowH = layout?.rowHeights?.[curR] ?? 20
    if (remH <= rowH) break
    remH -= rowH
    curR++
  }
  const toRow = curR
  const toRowOff = Math.round(remH * 9525)

  return { toCol, toColOff, toRow, toRowOff }
}

/** Generates <xdr:twoCellAnchor> XML for a chart. */
export function buildTwoCellAnchorXml(
  spec: ChartSpec,
  chartRelId: string,
  index: number,
  layout?: Partial<SheetLayout>
): string {
  const fromCol = spec.anchor.c
  const fromColOff = Math.round(spec.anchor.dx * 9525)
  const fromRow = spec.anchor.r
  const fromRowOff = Math.round(spec.anchor.dy * 9525)

  const { toCol, toColOff, toRow, toRowOff } = computeToAnchor(spec.anchor, spec.width, spec.height, layout)
  const frameId = index + 100

  return `<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from>
    <xdr:col>${fromCol}</xdr:col>
    <xdr:colOff>${fromColOff}</xdr:colOff>
    <xdr:row>${fromRow}</xdr:row>
    <xdr:rowOff>${fromRowOff}</xdr:rowOff>
  </xdr:from>
  <xdr:to>
    <xdr:col>${toCol}</xdr:col>
    <xdr:colOff>${toColOff}</xdr:colOff>
    <xdr:row>${toRow}</xdr:row>
    <xdr:rowOff>${toRowOff}</xdr:rowOff>
  </xdr:to>
  <xdr:graphicFrame macro="">
    <xdr:nvGraphicFramePr>
      <xdr:cNvPr id="${frameId}" name="Chart ${index + 1}"/>
      <xdr:cNvGraphicFramePr/>
    </xdr:nvGraphicFramePr>
    <xdr:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
    </xdr:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${chartRelId}"/>
      </a:graphicData>
    </a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>`
}

/** Parses cell reference or range from a formula string like `Sheet1!$A$1:$B$10`. */
function parseFormulaRange(f: string): Range | null {
  const stripped = f.replace(/^.*?!(?=\$?[A-Z])/, '').replace(/\$/g, '')
  if (stripped.includes(':')) {
    const parts = stripped.split(':')
    const m1 = /^([A-Z]+)(\d+)$/.exec(parts[0])
    const m2 = /^([A-Z]+)(\d+)$/.exec(parts[1])
    if (!m1 || !m2) return null
    const colIdx = (name: string): number => {
      let n = 0
      for (const ch of name) n = n * 26 + (ch.charCodeAt(0) - 64)
      return n - 1
    }
    return {
      r1: +m1[2] - 1,
      c1: colIdx(m1[1]),
      r2: +m2[2] - 1,
      c2: colIdx(m2[1])
    }
  }
  const m = /^([A-Z]+)(\d+)$/.exec(stripped)
  if (!m) return null
  let n = 0
  for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64)
  const c = n - 1
  const r = +m[2] - 1
  return { r1: r, c1: c, r2: r, c2: c }
}

/** Parses drawingml charts for a worksheet part. */
export function parseDrawingMlCharts(
  files: Record<string, Uint8Array>,
  dec: TextDecoder,
  sheetPath: string,
  sheetXml: string,
  colWidths: Record<number, number> = {},
  rowHeights: Record<number, number> = {}
): ChartSpec[] {
  const m = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)
  if (!m) return []
  const drawingRId = m[1]

  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const file = sheetPath.split('/').pop() ?? ''
  const sheetRelsPath = dir + '/_rels/' + file + '.rels'
  const sheetRelsBytes = files[sheetRelsPath]
  if (!sheetRelsBytes) return []

  const sheetRels = parseRelationships(dec.decode(sheetRelsBytes))
  const drawingRel = sheetRels.find((r) => r.id === drawingRId)
  if (!drawingRel) return []

  const drawingPath = resolveRelTarget(sheetPath, drawingRel.target)
  const drawingBytes = files[drawingPath]
  if (!drawingBytes) return []

  const drawingXml = dec.decode(drawingBytes)
  const drawingDir = drawingPath.split('/').slice(0, -1).join('/')
  const drawingFile = drawingPath.split('/').pop() ?? ''
  const drawingRelsPath = drawingDir + '/_rels/' + drawingFile + '.rels'
  const drawingRelsBytes = files[drawingRelsPath]
  const drawingRelMap = new Map<string, string>()
  if (drawingRelsBytes) {
    for (const r of parseRelationships(dec.decode(drawingRelsBytes))) {
      drawingRelMap.set(r.id, r.target)
    }
  }

  const out: ChartSpec[] = []
  const anchorRe = /<(?:xdr:)?twoCellAnchor\b([\s\S]*?)<\/(?:xdr:)?twoCellAnchor>/g
  let aMatch: RegExpExecArray | null

  while ((aMatch = anchorRe.exec(drawingXml))) {
    const block = aMatch[1]
    const chartRelM = /<(?:c:)?chart\b[^>]*r:id="([^"]+)"/.exec(block)
    if (!chartRelM) continue
    const chartRId = chartRelM[1]
    const chartTarget = drawingRelMap.get(chartRId)
    if (!chartTarget) continue

    const chartPath = resolveRelTarget(drawingPath, chartTarget)
    const chartBytes = files[chartPath]
    if (!chartBytes) continue

    const chartXml = dec.decode(chartBytes)

    // Parse Anchor coordinates
    const fromColM = /<(?:xdr:)?col>(\d+)<\/(?:xdr:)?col>/.exec(block)
    const fromColOffM = /<(?:xdr:)?colOff>(\d+)<\/(?:xdr:)?colOff>/.exec(block)
    const fromRowM = /<(?:xdr:)?row>(\d+)<\/(?:xdr:)?row>/.exec(block)
    const fromRowOffM = /<(?:xdr:)?rowOff>(\d+)<\/(?:xdr:)?rowOff>/.exec(block)

    const toColM = /<(?:xdr:)?to>[\s\S]*?<(?:xdr:)?col>(\d+)<\/(?:xdr:)?col>/.exec(block)
    const toColOffM = /<(?:xdr:)?to>[\s\S]*?<(?:xdr:)?colOff>(\d+)<\/(?:xdr:)?colOff>/.exec(block)
    const toRowM = /<(?:xdr:)?to>[\s\S]*?<(?:xdr:)?row>(\d+)<\/(?:xdr:)?row>/.exec(block)
    const toRowOffM = /<(?:xdr:)?to>[\s\S]*?<(?:xdr:)?rowOff>(\d+)<\/(?:xdr:)?rowOff>/.exec(block)

    const fromC = fromColM ? parseInt(fromColM[1], 10) : 0
    const fromR = fromRowM ? parseInt(fromRowM[1], 10) : 0
    const fromDx = fromColOffM ? Math.round(parseInt(fromColOffM[1], 10) / 9525) : 0
    const fromDy = fromRowOffM ? Math.round(parseInt(fromRowOffM[1], 10) / 9525) : 0

    const toC = toColM ? parseInt(toColM[1], 10) : fromC + 5
    const toR = toRowM ? parseInt(toRowM[1], 10) : fromR + 10
    const toDx = toColOffM ? Math.round(parseInt(toColOffM[1], 10) / 9525) : 0
    const toDy = toRowOffM ? Math.round(parseInt(toRowOffM[1], 10) / 9525) : 0

    // Compute width & height in pixels
    let width = 0
    if (toC === fromC) {
      width = Math.max(10, toDx - fromDx)
    } else {
      width = (colWidths[fromC] ?? 64) - fromDx
      for (let c = fromC + 1; c < toC; c++) width += colWidths[c] ?? 64
      width += toDx
    }

    let height = 0
    if (toR === fromR) {
      height = Math.max(10, toDy - fromDy)
    } else {
      height = (rowHeights[fromR] ?? 20) - fromDy
      for (let r = fromR + 1; r < toR; r++) height += rowHeights[r] ?? 20
      height += toDy
    }

    if (width <= 0) width = 480
    if (height <= 0) height = 288

    // Parse Title
    let title = ''
    let hasTitle = false
    const titleBlock = /<c:title>([\s\S]*?)<\/c:title>/.exec(chartXml)
    if (titleBlock) {
      const texts: string[] = []
      const tRe = /<[a-z0-9:]*t>([\s\S]*?)<\/[a-z0-9:]*t>/g
      let tM: RegExpExecArray | null
      while ((tM = tRe.exec(titleBlock[1]))) {
        texts.push(decodeXmlEntities(tM[1]))
      }
      title = texts.join('')
      hasTitle = !/<c:autoTitleDeleted\b[^>]*val="1"/.test(chartXml)
    }

    // Parse Legend
    let legendPosition: LegendPosition = 'none'
    let hasLegend = false
    const legendBlock = /<c:legend>([\s\S]*?)<\/c:legend>/.exec(chartXml)
    if (legendBlock) {
      hasLegend = true
      const posM = /<c:legendPos\b[^>]*val="([^"]+)"/.exec(legendBlock[1])
      const posVal = posM ? posM[1] : 'r'
      if (posVal === 't') legendPosition = 'top'
      else if (posVal === 'b') legendPosition = 'bottom'
      else if (posVal === 'l') legendPosition = 'left'
      else legendPosition = 'right'
    }

    const hasGridlines = /<c:majorGridlines\b/.test(chartXml)
    const hasDataLabels = /<c:dLbls\b/.test(chartXml)
    const hasAxes = /<c:(?:catAx|valAx)\b/.test(chartXml)

    // Determine Type & Subtype
    let type: ChartType = 'bar'
    let subtype: ChartSubtype = 'clusteredColumn'

    if (/<c:barChart\b/.test(chartXml)) {
      type = 'bar'
      const barDirM = /<c:barDir\b[^>]*val="([^"]+)"/.exec(chartXml)
      const groupM = /<c:grouping\b[^>]*val="([^"]+)"/.exec(chartXml)
      const barDir = barDirM ? barDirM[1] : 'col'
      const grouping = groupM ? groupM[1] : 'clustered'

      if (barDir === 'bar') {
        subtype = grouping === 'stacked' ? 'stackedBar' : grouping === 'percentStacked' ? 'percentStackedBar' : 'clusteredBar'
      } else {
        subtype = grouping === 'stacked' ? 'stackedColumn' : grouping === 'percentStacked' ? 'percentStackedColumn' : 'clusteredColumn'
      }
    } else if (/<c:lineChart\b/.test(chartXml)) {
      type = 'line'
      const groupM = /<c:grouping\b[^>]*val="([^"]+)"/.exec(chartXml)
      const grouping = groupM ? groupM[1] : 'standard'
      const hasMarkers = /<c:marker\b/.test(chartXml) && !/<c:symbol\b[^>]*val="none"/.test(chartXml)
      subtype = hasMarkers ? 'lineMarkers' : grouping === 'stacked' ? 'stackedLine' : 'line'
    } else if (/<c:areaChart\b/.test(chartXml)) {
      type = 'line'
      const groupM = /<c:grouping\b[^>]*val="([^"]+)"/.exec(chartXml)
      subtype = groupM && groupM[1] === 'stacked' ? 'stackedArea' : 'area'
    } else if (/<c:pieChart\b/.test(chartXml)) {
      type = 'pie'
      subtype = 'pie'
    } else if (/<c:doughnutChart\b/.test(chartXml)) {
      type = 'pie'
      subtype = 'doughnut'
    } else if (/<c:scatterChart\b/.test(chartXml)) {
      type = 'scatter'
      const styleM = /<c:scatterStyle\b[^>]*val="([^"]+)"/.exec(chartXml)
      subtype = styleM && styleM[1].includes('smooth') ? 'scatterSmooth' : 'scatter'
    }

    // Parse Series and Data Range
    const ranges: Range[] = []
    let detectedSeriesInColumns = true
    let firstValRange: Range | null = null

    const fRe = /<c:f>([\s\S]*?)<\/c:f>/g
    let fMatch: RegExpExecArray | null
    while ((fMatch = fRe.exec(chartXml))) {
      const formula = decodeXmlEntities(fMatch[1])
      const r = parseFormulaRange(formula)
      if (r) {
        ranges.push(r)
        if (!firstValRange && (chartXml.lastIndexOf('<c:val>', fMatch.index) > chartXml.lastIndexOf('</c:val>', fMatch.index) || chartXml.lastIndexOf('<c:yVal>', fMatch.index) > chartXml.lastIndexOf('</c:yVal>', fMatch.index))) {
          firstValRange = r
        }
      }
    }

    let dataRange: Range = { r1: 0, c1: 0, r2: 5, c2: 2 }
    if (ranges.length > 0) {
      let minR = Infinity
      let minC = Infinity
      let maxR = -Infinity
      let maxC = -Infinity
      for (const rg of ranges) {
        const n = normRange(rg)
        if (n.r1 < minR) minR = n.r1
        if (n.c1 < minC) minC = n.c1
        if (n.r2 > maxR) maxR = n.r2
        if (n.c2 > maxC) maxC = n.c2
      }
      dataRange = { r1: minR, c1: minC, r2: maxR, c2: maxC }
    }

    if (firstValRange) {
      const n = normRange(firstValRange)
      if (n.r2 > n.r1 && n.c1 === n.c2) {
        detectedSeriesInColumns = true
      } else if (n.c2 > n.c1 && n.r1 === n.r2) {
        detectedSeriesInColumns = false
      }
    }

    out.push({
      id: 'chart_' + Math.random().toString(36).slice(2, 9),
      type,
      subtype,
      dataRange,
      seriesInColumns: detectedSeriesInColumns,
      title,
      anchor: { r: fromR, c: fromC, dx: fromDx, dy: fromDy },
      width,
      height,
      legendPosition,
      elements: {
        title: hasTitle,
        legend: hasLegend,
        gridlines: hasGridlines,
        dataLabels: hasDataLabels,
        axes: hasAxes,
        axisTitles: false
      }
    })
  }

  return out
}

// Chart round-trip: write -> fast read, for every ChartSubtype, plus structural validation of the
// written chart/drawing XML (required elements present, in schema order) directly on the zip bytes.
// Deliberately does NOT go through fastXlsx.test.ts's expectEqualReads helper: readXlsxExceljs (the
// differential test's other half) has no chart support at all, so a workbook with charts would fail
// that comparison for reasons unrelated to charts.
import { describe, it, expect } from 'vitest'
import { unzipSync } from 'fflate'
import { writeXlsx, readXlsx } from './xlsx'
import type { WorkbookData } from '../model/store'
import type { ChartSpec, ChartSubtype, ChartType } from '../model/types'
import { defaultLayout } from '../model/defaults'

function cellsFromRows(rows: (string | number)[][]): Record<string, { raw?: string }> {
  const cells: Record<string, { raw?: string }> = {}
  rows.forEach((row, r) => {
    row.forEach((v, c) => {
      cells[r + ',' + c] = { raw: String(v) }
    })
  })
  return cells
}

const DATA_ROWS: (string | number)[][] = [
  ['', 'Q1', 'Q2'],
  ['North', 10, 20],
  ['South', 15, 25],
  ['East', 5, 12]
]

function makeChart(type: ChartType, subtype: ChartSubtype, overrides?: Partial<ChartSpec>): ChartSpec {
  const isPie = type === 'pie'
  return {
    id: 'chart1',
    type,
    subtype,
    dataRange: { r1: 0, c1: 0, r2: 3, c2: 2 },
    seriesInColumns: true,
    title: 'My Chart',
    anchor: { r: 6, c: 0, dx: 0, dy: 0 },
    width: 480,
    height: 288,
    legendPosition: isPie ? 'right' : 'bottom',
    elements: {
      title: true,
      legend: true,
      gridlines: !isPie,
      dataLabels: isPie,
      axes: !isPie,
      axisTitles: false
    },
    ...overrides
  }
}

function makeWorkbook(chart: ChartSpec): WorkbookData {
  return {
    sheets: [
      {
        name: 'Sheet1',
        cells: cellsFromRows(DATA_ROWS),
        layout: { ...defaultLayout(), charts: [chart] }
      }
    ],
    activeSheet: 0
  }
}

const SUBTYPES: { type: ChartType; subtype: ChartSubtype }[] = [
  { type: 'bar', subtype: 'clusteredColumn' },
  { type: 'bar', subtype: 'stackedColumn' },
  { type: 'bar', subtype: 'percentStackedColumn' },
  { type: 'bar', subtype: 'clusteredBar' },
  { type: 'bar', subtype: 'stackedBar' },
  { type: 'bar', subtype: 'percentStackedBar' },
  { type: 'line', subtype: 'line' },
  { type: 'line', subtype: 'stackedLine' },
  { type: 'line', subtype: 'lineMarkers' },
  { type: 'line', subtype: 'area' },
  { type: 'line', subtype: 'stackedArea' },
  { type: 'pie', subtype: 'pie' },
  { type: 'pie', subtype: 'doughnut' },
  { type: 'scatter', subtype: 'scatter' },
  { type: 'scatter', subtype: 'scatterSmooth' }
]

describe('chart write -> fast read round-trip', () => {
  for (const { type, subtype } of SUBTYPES) {
    it(`round-trips a ${subtype} chart`, async () => {
      const chart = makeChart(type, subtype)
      const wb = makeWorkbook(chart)
      const bytes = await writeXlsx(wb)
      const read = await readXlsx(bytes)

      const charts = read.sheets[0].layout?.charts
      expect(charts).toBeDefined()
      expect(charts).toHaveLength(1)
      const back = charts![0]

      expect(back.type).toBe(type)
      expect(back.subtype).toBe(subtype)
      expect(back.title).toBe('My Chart')
      expect(back.elements.title).toBe(true)
      expect(back.elements.legend).toBe(true)
      expect(back.legendPosition).toBe(chart.legendPosition)
      expect(back.elements.gridlines).toBe(chart.elements.gridlines)
      expect(back.elements.dataLabels).toBe(chart.elements.dataLabels)
      expect(back.elements.axes).toBe(chart.elements.axes)

      // Anchor position always round-trips exactly (whole px stored, no sub-px rounding).
      expect(back.anchor).toEqual(chart.anchor)
      // Width/height go through an EMU<->px + column/row-walk round trip; allow a few px of slack.
      expect(Math.abs(back.width - chart.width)).toBeLessThanOrEqual(3)
      expect(Math.abs(back.height - chart.height)).toBeLessThanOrEqual(3)

      // dataRange should recover the full A1:C4 block (header row + header column included, since
      // the series-name/category formulas point at them) - except pie/doughnut, which (like Excel
      // itself) only ever plots one series, so the block shrinks to that series' own column (B).
      const expectedRange = type === 'pie' ? { r1: 0, c1: 0, r2: 3, c2: 1 } : chart.dataRange
      expect(back.dataRange).toEqual(expectedRange)
      expect(back.seriesInColumns).toBe(true)
    })
  }

  it('round-trips a chart with no header row/column (positional categories, numLit fallback)', async () => {
    const chart = makeChart('bar', 'clusteredColumn', { dataRange: { r1: 0, c1: 0, r2: 2, c2: 1 }, title: '' , elements: { title: false, legend: false, gridlines: true, dataLabels: false, axes: true, axisTitles: false } })
    const wb: WorkbookData = {
      sheets: [
        {
          name: 'Sheet1',
          cells: cellsFromRows([[1, 2], [3, 4], [5, 6]]),
          layout: { ...defaultLayout(), charts: [chart] }
        }
      ],
      activeSheet: 0
    }
    const bytes = await writeXlsx(wb)
    const read = await readXlsx(bytes)
    const back = read.sheets[0].layout?.charts?.[0]
    expect(back).toBeDefined()
    expect(back!.elements.title).toBe(false)
    expect(back!.elements.legend).toBe(false)
  })

  it('round-trips axis titles', async () => {
    const chart = makeChart('bar', 'clusteredColumn', {
      elements: { title: true, legend: true, gridlines: true, dataLabels: false, axes: true, axisTitles: true },
      axisTitles: { x: 'Region', y: 'Amount' }
    })
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const read = await readXlsx(bytes)
    const back = read.sheets[0].layout?.charts?.[0]
    expect(back?.elements.axisTitles).toBe(true)
    expect(back?.axisTitles?.x).toBe('Region')
    expect(back?.axisTitles?.y).toBe('Amount')
  })

  it('round-trips legend position none (no legend element written)', async () => {
    const chart = makeChart('line', 'line', { legendPosition: 'none', elements: { title: true, legend: false, gridlines: true, dataLabels: false, axes: true, axisTitles: false } })
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const read = await readXlsx(bytes)
    const back = read.sheets[0].layout?.charts?.[0]
    expect(back?.elements.legend).toBe(false)
    expect(back?.legendPosition).toBe('none')
  })

  it('writes distinct chart parts and anchors for multiple charts on one sheet', async () => {
    const chart1 = makeChart('bar', 'clusteredColumn', { id: 'a' })
    const chart2 = makeChart('pie', 'pie', { id: 'b', anchor: { r: 20, c: 0, dx: 0, dy: 0 } })
    const wb: WorkbookData = {
      sheets: [{ name: 'Sheet1', cells: cellsFromRows(DATA_ROWS), layout: { ...defaultLayout(), charts: [chart1, chart2] } }],
      activeSheet: 0
    }
    const bytes = await writeXlsx(wb)
    const read = await readXlsx(bytes)
    const charts = read.sheets[0].layout?.charts
    expect(charts).toHaveLength(2)
    expect(new Set(charts!.map((c) => c.type))).toEqual(new Set(['bar', 'pie']))

    const files = unzipSync(bytes)
    expect(files['xl/charts/chart1.xml']).toBeDefined()
    expect(files['xl/charts/chart2.xml']).toBeDefined()
  })

  it('places charts on the right sheet when only a later sheet has one', async () => {
    const chart = makeChart('bar', 'clusteredColumn')
    const wb: WorkbookData = {
      sheets: [
        { name: 'Plain', cells: cellsFromRows([['x']]), layout: defaultLayout() },
        { name: 'Charted', cells: cellsFromRows(DATA_ROWS), layout: { ...defaultLayout(), charts: [chart] } }
      ],
      activeSheet: 0
    }
    const bytes = await writeXlsx(wb)
    const read = await readXlsx(bytes)
    expect(read.sheets[0].layout?.charts ?? []).toHaveLength(0)
    expect(read.sheets[1].layout?.charts).toHaveLength(1)
  })
})

// ---------- structural XML validation ----------
// Confirms the *raw written bytes* contain schema-valid, correctly-ordered chart/drawing XML -
// not just that our own reader can parse it back (round-trip tests above could pass even if both
// sides shared the same bug). Checks presence and relative order of the elements OOXML's chart
// schema (CT_Chart/CT_PlotArea/CT_BarChart/etc) requires, using indexOf position comparisons.
function assertOrder(xml: string, tagsInOrder: string[]): void {
  let pos = -1
  for (const tag of tagsInOrder) {
    const idx = xml.indexOf(tag, pos + 1)
    expect(idx, `expected to find "${tag}" after position ${pos} in:\n${xml}`).toBeGreaterThan(pos)
    pos = idx
  }
}

describe('written chart XML structure', () => {
  it('[Content_Types].xml declares the drawing and chart part content types', async () => {
    const chart = makeChart('bar', 'clusteredColumn')
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const ct = new TextDecoder().decode(files['[Content_Types].xml'])
    expect(ct).toContain('PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"')
    expect(ct).toContain('PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"')
  })

  it('sheet1.xml has a <drawing r:id> element and a matching relationship', async () => {
    const chart = makeChart('bar', 'clusteredColumn')
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const sheetXml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    const m = /<drawing r:id="(rId\d+)"\/>/.exec(sheetXml)
    expect(m).not.toBeNull()
    // <drawing> must appear inside <worksheet>...</worksheet>, and before the closing tag.
    expect(sheetXml.indexOf(m![0])).toBeLessThan(sheetXml.indexOf('</worksheet>'))

    const relsXml = new TextDecoder().decode(files['xl/worksheets/_rels/sheet1.xml.rels'])
    expect(relsXml).toContain(`Id="${m![1]}"`)
    expect(relsXml).toContain('Target="../drawings/drawing1.xml"')
    expect(relsXml).toContain('relationships/drawing')
  })

  it('drawing1.xml has a schema-ordered twoCellAnchor -> graphicFrame -> chart reference', async () => {
    const chart = makeChart('bar', 'clusteredColumn')
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const xml = new TextDecoder().decode(files['xl/drawings/drawing1.xml'])
    assertOrder(xml, [
      '<xdr:wsDr',
      '<xdr:twoCellAnchor>',
      '<xdr:from>',
      '</xdr:from>',
      '<xdr:to>',
      '</xdr:to>',
      '<xdr:graphicFrame',
      '<xdr:nvGraphicFramePr>',
      '</xdr:nvGraphicFramePr>',
      '<xdr:xfrm>',
      '<a:graphic>',
      '<a:graphicData',
      '<c:chart',
      'r:id="rId1"',
      '</a:graphicData>',
      '</a:graphic>',
      '</xdr:graphicFrame>',
      '<xdr:clientData/>',
      '</xdr:twoCellAnchor>',
      '</xdr:wsDr>'
    ])

    const relsXml = new TextDecoder().decode(files['xl/drawings/_rels/drawing1.xml.rels'])
    expect(relsXml).toContain('Id="rId1"')
    expect(relsXml).toContain('Target="../charts/chart1.xml"')
  })

  it('a bar chart part follows CT_Chart/CT_PlotArea/CT_BarChart/CT_BarSer schema order', async () => {
    const chart = makeChart('bar', 'stackedColumn', {
      elements: { title: true, legend: true, gridlines: true, dataLabels: true, axes: true, axisTitles: true },
      axisTitles: { x: 'Region', y: 'Amount' }
    })
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const xml = new TextDecoder().decode(files['xl/charts/chart1.xml'])

    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true)
    assertOrder(xml, [
      '<c:chartSpace',
      '<c:chart>',
      '<c:title>',
      '</c:title>',
      '<c:plotArea>',
      '<c:layout/>',
      '<c:barChart>',
      '<c:barDir val="col"/>',
      '<c:grouping val="stacked"/>',
      '<c:varyColors val="0"/>',
      '<c:ser>',
      '<c:idx val="0"/>',
      '<c:order val="0"/>',
      '<c:tx>',
      '<c:spPr>',
      '<c:cat>',
      '<c:val>',
      '</c:ser>',
      '<c:dLbls>',
      '<c:gapWidth val="150"/>',
      '<c:overlap val="100"/>',
      '<c:axId val="111111111"/>',
      '<c:axId val="222222222"/>',
      '</c:barChart>',
      '<c:catAx>',
      '<c:axId val="111111111"/>',
      '<c:scaling>',
      '<c:delete val="0"/>',
      '<c:axPos val="b"/>',
      '<c:title>',
      '<c:crossAx val="222222222"/>',
      '</c:catAx>',
      '<c:valAx>',
      '<c:axId val="222222222"/>',
      '<c:majorGridlines/>',
      '<c:title>',
      '<c:crossAx val="111111111"/>',
      '</c:valAx>',
      '</c:plotArea>',
      '<c:legend>',
      '<c:legendPos val="b"/>',
      '</c:legend>',
      '<c:plotVisOnly val="1"/>',
      '</c:chart>',
      '</c:chartSpace>'
    ])

    // Cached values are present for every literal (non-formula) data cell.
    expect(xml).toContain('<c:v>10</c:v>')
    expect(xml).toContain('<c:v>North</c:v>')
  })

  it('a pie chart part has per-slice dPt colors and no axes', async () => {
    const chart = makeChart('pie', 'pie')
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const xml = new TextDecoder().decode(files['xl/charts/chart1.xml'])
    assertOrder(xml, ['<c:pieChart>', '<c:varyColors val="1"/>', '<c:ser>', '<c:dPt>', '</c:ser>', '<c:firstSliceAng val="0"/>', '</c:pieChart>'])
    expect(xml).not.toContain('<c:catAx>')
    expect(xml).not.toContain('<c:valAx>')
  })

  it('a doughnut chart part has a holeSize after firstSliceAng', async () => {
    const chart = makeChart('pie', 'doughnut')
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const xml = new TextDecoder().decode(files['xl/charts/chart1.xml'])
    assertOrder(xml, ['<c:doughnutChart>', '<c:firstSliceAng val="0"/>', '<c:holeSize val="50"/>', '</c:doughnutChart>'])
  })

  it('a scatter chart part uses xVal/yVal and two valAx elements', async () => {
    const chart = makeChart('scatter', 'scatterSmooth', { dataRange: { r1: 0, c1: 0, r2: 3, c2: 2 } })
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const files = unzipSync(bytes)
    const xml = new TextDecoder().decode(files['xl/charts/chart1.xml'])
    assertOrder(xml, [
      '<c:scatterChart>',
      '<c:scatterStyle val="smoothMarker"/>',
      '<c:ser>',
      '<c:marker>',
      '<c:xVal>',
      '<c:yVal>',
      '<c:smooth val="1"/>',
      '</c:ser>',
      '<c:axId val="333333333"/>',
      '<c:axId val="444444444"/>',
      '</c:scatterChart>',
      '<c:valAx>',
      '<c:axId val="333333333"/>',
      '<c:axPos val="b"/>',
      '</c:valAx>',
      '<c:valAx>',
      '<c:axId val="444444444"/>',
      '<c:axPos val="l"/>',
      '</c:valAx>'
    ])
  })

  it('a line-with-markers chart part has visible markers; a plain line chart marks symbol none', async () => {
    const markersWb = makeWorkbook(makeChart('line', 'lineMarkers'))
    const markersXml = new TextDecoder().decode(unzipSync(await writeXlsx(markersWb))['xl/charts/chart1.xml'])
    expect(markersXml).toContain('<c:symbol val="circle"/>')

    const plainWb = makeWorkbook(makeChart('line', 'line'))
    const plainXml = new TextDecoder().decode(unzipSync(await writeXlsx(plainWb))['xl/charts/chart1.xml'])
    expect(plainXml).toContain('<c:symbol val="none"/>')
  })

  it('series colors follow the Office theme accent rotation for colorSetId 0', async () => {
    const chart = makeChart('bar', 'clusteredColumn', { colorSetId: 0 })
    const wb = makeWorkbook(chart)
    const bytes = await writeXlsx(wb)
    const xml = new TextDecoder().decode(unzipSync(bytes)['xl/charts/chart1.xml'])
    expect(xml).toContain('<a:srgbClr val="4472C4"/>') // accent1
    expect(xml).toContain('<a:srgbClr val="ED7D31"/>') // accent2
  })
})

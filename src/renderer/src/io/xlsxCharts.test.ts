import { describe, expect, it } from 'vitest'
import { buildChartSpaceXml, parseComboSeries } from './xlsxCharts'
import { createChartSpec, COMBO_GALLERY, WATERFALL_ENTRY, RADAR_GALLERY, STOCK_GALLERY, SURFACE_GALLERY, HIERARCHY_GALLERY, STATISTIC_GALLERY, FUNNEL_ENTRY } from '../charts/gallery'

const grid = [
  ['', 'Sales', 'Margin'],
  ['Q1', 10, 2],
  ['Q2', -4, 3],
  ['Q3', 8, 1]
]
const reader = { get: (r: number, c: number) => (grid[r]?.[c] ?? null) as never }
const range = { r1: 0, c1: 0, r2: 3, c2: 2 }
const anchor = { r: 5, c: 0, dx: 0, dy: 0 }

describe('combo chart xlsx', () => {
  it('writes bar + line groups with a secondary axis and reads them back', () => {
    const spec = createChartSpec(COMBO_GALLERY[1], range, anchor)
    spec.comboSeries = [{ type: 'bar' }, { type: 'line', secondary: true }]
    const xml = buildChartSpaceXml(spec, 'Sheet1', reader)
    expect(xml).toContain('<c:barChart>')
    expect(xml).toContain('<c:lineChart>')
    expect(xml).toContain('<c:crosses val="max"/>')
    expect(parseComboSeries(xml)).toEqual([{ type: 'bar' }, { type: 'line', secondary: true }])
  })

  it('does not treat a plain bar chart as combo', () => {
    expect(parseComboSeries('<c:barChart><c:ser><c:idx val="0"/></c:ser><c:axId val="1"/></c:barChart>')).toBeNull()
  })
})

describe('waterfall chart xlsx', () => {
  it('writes a stacked column with an invisible base and hides its legend entry', () => {
    const spec = createChartSpec(WATERFALL_ENTRY, range, anchor)
    const xml = buildChartSpaceXml(spec, 'Sheet1', reader)
    expect(xml).toContain('<c:grouping val="stacked"/>')
    expect(xml).toContain('<a:noFill/>')
    expect(xml).toContain('<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>')
  })
})

describe('extra chart families xlsx', () => {
  it('writes radar and stock charts natively and every family produces a plot', () => {
    for (const [e, tag] of [[RADAR_GALLERY[2], '<c:radarStyle val="filled"/>'], [STOCK_GALLERY[0], '<c:hiLowLines/>'], [STOCK_GALLERY[1], '<c:upDownBars>'], [SURFACE_GALLERY[0], '<c:lineChart>'], [HIERARCHY_GALLERY[0], '<c:barChart>'], [STATISTIC_GALLERY[0], '<c:gapWidth val="5"/>'], [STATISTIC_GALLERY[2], '<c:barChart>'], [FUNNEL_ENTRY, '<c:barChart>']] as const) {
      const xml = buildChartSpaceXml(createChartSpec(e, range, anchor), 'Sheet1', reader)
      expect(xml, e.label).toContain(tag)
      expect(xml, e.label).toContain('<c:axId')
    }
  })
})

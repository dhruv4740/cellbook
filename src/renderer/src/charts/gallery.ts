// Chart-type galleries shown under the Insert tab's split buttons, and the factory that turns a
// picked gallery entry + source range into a ChartSpec with Excel's default elements.
import { ChartSpec, ChartSubtype, ChartType, Range } from '../model/types'
import { defaultSeriesInColumns } from './dataExtraction'

export interface GalleryEntry {
  subtype: ChartSubtype
  type: ChartType
  label: string
  /** Combo entries: put the line series on a secondary value axis. */
  comboSecondary?: boolean
}

export const COMBO_GALLERY: GalleryEntry[] = [
  { subtype: 'combo', type: 'combo', label: 'Clustered Column - Line' },
  { subtype: 'combo', type: 'combo', label: 'Clustered Column - Line on Secondary Axis', comboSecondary: true }
]

export const WATERFALL_ENTRY: GalleryEntry = { subtype: 'waterfall', type: 'waterfall', label: 'Waterfall' }

export const COLUMN_BAR_GALLERY: GalleryEntry[] = [
  { subtype: 'clusteredColumn', type: 'bar', label: 'Clustered Column' },
  { subtype: 'stackedColumn', type: 'bar', label: 'Stacked Column' },
  { subtype: 'percentStackedColumn', type: 'bar', label: '100% Stacked Column' },
  { subtype: 'clusteredBar', type: 'bar', label: 'Clustered Bar' },
  { subtype: 'stackedBar', type: 'bar', label: 'Stacked Bar' },
  { subtype: 'percentStackedBar', type: 'bar', label: '100% Stacked Bar' }
]

export const LINE_AREA_GALLERY: GalleryEntry[] = [
  { subtype: 'line', type: 'line', label: 'Line' },
  { subtype: 'stackedLine', type: 'line', label: 'Stacked Line' },
  { subtype: 'lineMarkers', type: 'line', label: 'Line with Markers' },
  { subtype: 'area', type: 'line', label: 'Area' },
  { subtype: 'stackedArea', type: 'line', label: 'Stacked Area' }
]

export const PIE_DOUGHNUT_GALLERY: GalleryEntry[] = [
  { subtype: 'pie', type: 'pie', label: 'Pie' },
  { subtype: 'doughnut', type: 'pie', label: 'Doughnut' }
]

export const SCATTER_BUBBLE_GALLERY: GalleryEntry[] = [
  { subtype: 'scatter', type: 'scatter', label: 'Scatter' },
  { subtype: 'scatterSmooth', type: 'scatter', label: 'Scatter with Smooth Lines' }
]

export const ALL_GALLERIES = [
  { name: 'Column or Bar', entries: COLUMN_BAR_GALLERY },
  { name: 'Line or Area', entries: LINE_AREA_GALLERY },
  { name: 'Pie or Doughnut', entries: PIE_DOUGHNUT_GALLERY },
  { name: 'Scatter (X, Y) or Bubble', entries: SCATTER_BUBBLE_GALLERY }
]

let nextChartId = 1
export function newChartId(): string {
  return 'chart' + Date.now().toString(36) + '_' + nextChartId++
}

/** Excel's default new-chart size: ~5" x 3" at 96dpi. */
export const DEFAULT_CHART_WIDTH = 480
export const DEFAULT_CHART_HEIGHT = 288

export function createChartSpec(
  entry: GalleryEntry,
  dataRange: Range,
  anchor: { r: number; c: number; dx: number; dy: number }
): ChartSpec {
  const isPie = entry.type === 'pie'
  return {
    id: newChartId(),
    type: entry.type,
    subtype: entry.subtype,
    dataRange,
    seriesInColumns: defaultSeriesInColumns(dataRange),
    title: 'Chart Title',
    anchor,
    width: DEFAULT_CHART_WIDTH,
    height: DEFAULT_CHART_HEIGHT,
    legendPosition: isPie ? 'right' : 'bottom',
    elements: {
      title: true,
      legend: true,
      gridlines: !isPie,
      dataLabels: isPie,
      axes: !isPie,
      axisTitles: false
    }
  }
}

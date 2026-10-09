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

export const RADAR_GALLERY: GalleryEntry[] = [
  { subtype: 'radar', type: 'radar', label: 'Radar' },
  { subtype: 'radarMarkers', type: 'radar', label: 'Radar with Markers' },
  { subtype: 'filledRadar', type: 'radar', label: 'Filled Radar' }
]
export const FUNNEL_ENTRY: GalleryEntry = { subtype: 'funnel', type: 'funnel', label: 'Funnel' }
export const STOCK_GALLERY: GalleryEntry[] = [
  { subtype: 'stockHLC', type: 'stock', label: 'High-Low-Close' },
  { subtype: 'stockOHLC', type: 'stock', label: 'Open-High-Low-Close' }
]
export const SURFACE_GALLERY: GalleryEntry[] = [
  { subtype: 'surface3D', type: 'surface', label: '3-D Surface' },
  { subtype: 'wireSurface', type: 'surface', label: 'Wireframe 3-D Surface' },
  { subtype: 'contour', type: 'surface', label: 'Contour' },
  { subtype: 'wireContour', type: 'surface', label: 'Wireframe Contour' }
]
export const HIERARCHY_GALLERY: GalleryEntry[] = [
  { subtype: 'treemap', type: 'hierarchy', label: 'Treemap' },
  { subtype: 'sunburst', type: 'hierarchy', label: 'Sunburst' }
]
export const STATISTIC_GALLERY: GalleryEntry[] = [
  { subtype: 'histogram', type: 'statistic', label: 'Histogram' },
  { subtype: 'pareto', type: 'statistic', label: 'Pareto' },
  { subtype: 'boxWhisker', type: 'statistic', label: 'Box and Whisker' }
]

/** Types whose legend Excel hides by default (one series coloured by point, or the body draws its own key). */
export const NO_LEGEND_TYPES: ChartType[] = ['funnel', 'surface', 'hierarchy', 'statistic']

export const ALL_GALLERIES = [
  { name: 'Column or Bar', entries: COLUMN_BAR_GALLERY },
  { name: 'Line or Area', entries: LINE_AREA_GALLERY },
  { name: 'Pie or Doughnut', entries: PIE_DOUGHNUT_GALLERY },
  { name: 'Scatter (X, Y) or Bubble', entries: SCATTER_BUBBLE_GALLERY },
  { name: 'Radar', entries: RADAR_GALLERY },
  { name: 'Stock', entries: STOCK_GALLERY },
  { name: 'Surface', entries: SURFACE_GALLERY },
  { name: 'Hierarchy', entries: HIERARCHY_GALLERY },
  { name: 'Statistic', entries: STATISTIC_GALLERY },
  { name: 'Funnel', entries: [FUNNEL_ENTRY] }
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

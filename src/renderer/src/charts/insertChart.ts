// Wires a picked gallery entry (from the Insert tab or Chart Design > Change Chart Type) to an
// actual chart on the active sheet, using Excel's current-region behavior for the source range.
import { WorkbookStore, SheetState } from '../model/store'
import { setUi } from '../chrome/ui'
import { ChartSpec, MAX_COLS, MAX_ROWS } from '../model/types'
import { resolveChartRange, extractChartData, defaultSeriesInColumns } from './dataExtraction'
import { readerFor, insertChart as insertChartSpec } from './chartsStore'
import { createChartSpec, GalleryEntry, DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT, NO_LEGEND_TYPES, COLUMN_BAR_GALLERY } from './gallery'
import { defaultComboSeries } from './chartCombo'

function sourceFromSelection(store: WorkbookStore) {
  const sheet: SheetState = store.active
  const sel = sheet.selection.active
  const selRange = sheet.selection.ranges[sheet.selection.ranges.length - 1] ?? { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }
  const reader = readerFor(store, sheet)
  return { sheet, reader, dataRange: resolveChartRange(reader, selRange, MAX_ROWS - 1, MAX_COLS - 1) }
}

/** Series names a chart inserted from the current selection would get (for the custom Combo dialog). */
export function selectionSeriesNames(store: WorkbookStore): string[] {
  const { reader, dataRange } = sourceFromSelection(store)
  return extractChartData(reader, dataRange, defaultSeriesInColumns(dataRange)).series.map((s) => s.name)
}

export function insertChartFromSelection(store: WorkbookStore, entry: GalleryEntry, comboSeries?: ChartSpec['comboSeries']): void {
  const { sheet, reader, dataRange } = sourceFromSelection(store)

  // Anchor a couple of rows below the data, like Excel drops a new chart near the selection,
  // clamped so it stays on-screen for a freshly opened sheet.
  const anchor = { r: Math.min(MAX_ROWS - 1, dataRange.r2 + 2), c: dataRange.c1, dx: 0, dy: 0 }
  const spec = createChartSpec(entry, dataRange, anchor)
  spec.width = DEFAULT_CHART_WIDTH
  spec.height = DEFAULT_CHART_HEIGHT
  // Excel titles a one-series chart with the series name, and drops the legend unless it is a
  // pie/doughnut (where the legend names the slices) or a waterfall (Increase / Decrease).
  const data = extractChartData(reader, dataRange, spec.seriesInColumns, spec.firstIsCategories)
  if (spec.type === 'combo') spec.comboSeries = comboSeries ?? defaultComboSeries(data.series.length, !!entry.comboSecondary)
  if (data.series.length === 1) {
    if (data.series[0].name) spec.title = data.series[0].name
    if (spec.type !== 'pie' && spec.type !== 'waterfall') spec.elements.legend = false
  }
  if (NO_LEGEND_TYPES.includes(spec.type)) { spec.elements.legend = false; spec.elements.gridlines = spec.type === 'statistic'; spec.elements.axes = spec.type === 'statistic' }
  insertChartSpec(store, sheet, spec)
  setUi({ selectedChart: spec.id, ribbonTab: 'Chart Design' })
}

/** Insert > PivotChart: a clustered column chart of the PivotTable under the active cell (Grand Total row/column left out).
 *  Without a PivotTable it opens the PivotTable dialog, like Excel's Create PivotChart. The chart plots the pivot's current
 *  output range; it is not re-linked when the pivot's size changes. */
export function insertPivotChart(store: WorkbookStore): void {
  const sheet = store.active
  const { r, c } = sheet.selection.active
  const pv = (sheet.layout.pivots ?? []).find((p) => p.lastSize && r >= p.anchor.r && r < p.anchor.r + p.lastSize.rows && c >= p.anchor.c && c < p.anchor.c + p.lastSize.cols)
  if (!pv || !pv.lastSize) { setUi({ dialog: { kind: 'pivot' } }); return }
  let r2 = pv.anchor.r + pv.lastSize.rows - 1, c2 = pv.anchor.c + pv.lastSize.cols - 1
  const text = (rr: number, cc: number): string => String(store.getValue(sheet, rr, cc) ?? '')
  if (/^grand total$/i.test(text(r2, pv.anchor.c))) r2--
  if (c2 > pv.anchor.c && /^grand total$/i.test(text(pv.anchor.r + (pv.cols.length ? 1 : 0), c2))) c2--
  const dataRange = { r1: pv.anchor.r + (pv.cols.length ? 1 : 0), c1: pv.anchor.c, r2, c2 }
  if (dataRange.r2 <= dataRange.r1 || dataRange.c2 <= dataRange.c1) return
  const reader = readerFor(store, sheet)
  const spec = createChartSpec(COLUMN_BAR_GALLERY[0], dataRange, { r: Math.min(MAX_ROWS - 1, r2 + 2), c: pv.anchor.c, dx: 0, dy: 0 })
  const data = extractChartData(reader, dataRange, spec.seriesInColumns, spec.firstIsCategories)
  if (data.series.length === 1 && data.series[0].name) { spec.title = data.series[0].name; spec.elements.legend = false }
  insertChartSpec(store, sheet, spec)
  setUi({ selectedChart: spec.id, ribbonTab: 'Chart Design' })
}

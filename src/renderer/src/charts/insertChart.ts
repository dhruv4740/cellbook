// Wires a picked gallery entry (from the Insert tab or Chart Design > Change Chart Type) to an
// actual chart on the active sheet, using Excel's current-region behavior for the source range.
import { WorkbookStore, SheetState } from '../model/store'
import { setUi } from '../chrome/ui'
import { MAX_COLS, MAX_ROWS } from '../model/types'
import { resolveChartRange } from './dataExtraction'
import { readerFor, insertChart as insertChartSpec } from './chartsStore'
import { createChartSpec, GalleryEntry, DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT } from './gallery'

export function insertChartFromSelection(store: WorkbookStore, entry: GalleryEntry): void {
  const sheet: SheetState = store.active
  const sel = sheet.selection.active
  const selRange = sheet.selection.ranges[sheet.selection.ranges.length - 1] ?? { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }
  const dataRange = resolveChartRange(readerFor(store, sheet), selRange, MAX_ROWS - 1, MAX_COLS - 1)

  // Anchor a couple of rows below the data, like Excel drops a new chart near the selection,
  // clamped so it stays on-screen for a freshly opened sheet.
  const anchor = { r: Math.min(MAX_ROWS - 1, dataRange.r2 + 2), c: dataRange.c1, dx: 0, dy: 0 }
  const spec = createChartSpec(entry, dataRange, anchor)
  spec.width = DEFAULT_CHART_WIDTH
  spec.height = DEFAULT_CHART_HEIGHT
  insertChartSpec(store, sheet, spec)
  setUi({ selectedChart: spec.id, ribbonTab: 'Chart Design' })
}

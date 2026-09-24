// Store-facing helpers for chart objects. Every mutation goes through store.updateLayout so it
// undoes/redoes like every other layout change; nothing here touches the store's internals directly.
import { WorkbookStore, SheetState } from '../model/store'
import { ChartSpec, Range } from '../model/types'
import { CellReader } from './dataExtraction'

/** Charts on a sheet, defaulting to []; layouts created before this feature won't have the field. */
export function chartsOf(sheet: SheetState): ChartSpec[] {
  return sheet.layout.charts ?? []
}

/** A CellReader backed by the live store — used to (re-)extract chart data on every render. */
export function readerFor(store: WorkbookStore, sheet: SheetState): CellReader {
  return { get: (r, c) => store.getValue(sheet, r, c) }
}

export function insertChart(store: WorkbookStore, sheet: SheetState, spec: ChartSpec): void {
  store.updateLayout(sheet, 'Insert Chart', (l) => {
    l.charts = [...(l.charts ?? []), spec]
  })
}

export function deleteChart(store: WorkbookStore, sheet: SheetState, id: string): void {
  store.updateLayout(sheet, 'Delete Chart', (l) => {
    l.charts = (l.charts ?? []).filter((c) => c.id !== id)
  })
}

export function updateChart(store: WorkbookStore, sheet: SheetState, id: string, patch: Partial<ChartSpec>, label = 'Format Chart'): void {
  store.updateLayout(sheet, label, (l) => {
    l.charts = (l.charts ?? []).map((c) => (c.id === id ? { ...c, ...patch } : c))
  })
}

export function moveOrResizeChart(store: WorkbookStore, sheet: SheetState, id: string, anchor: ChartSpec['anchor'], width: number, height: number): void {
  updateChart(store, sheet, id, { anchor, width, height }, 'Move or Resize Chart')
}

export function setChartRange(store: WorkbookStore, sheet: SheetState, id: string, dataRange: Range, seriesInColumns: boolean): void {
  updateChart(store, sheet, id, { dataRange, seriesInColumns }, 'Select Data')
}

export function bringChartToFront(store: WorkbookStore, sheet: SheetState, id: string): void {
  store.updateLayout(sheet, 'Bring to Front', (l) => {
    const list = l.charts ?? []
    const chart = list.find((c) => c.id === id)
    if (!chart) return
    l.charts = [...list.filter((c) => c.id !== id), chart]
  })
}

export function sendChartToBack(store: WorkbookStore, sheet: SheetState, id: string): void {
  store.updateLayout(sheet, 'Send to Back', (l) => {
    const list = l.charts ?? []
    const chart = list.find((c) => c.id === id)
    if (!chart) return
    l.charts = [chart, ...list.filter((c) => c.id !== id)]
  })
}

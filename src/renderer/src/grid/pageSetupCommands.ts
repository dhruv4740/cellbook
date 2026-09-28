// Page Layout commands: page setup edits (undoable), manual page breaks, scale to fit.
import { store } from '../model/store'
import { PageSetup } from '../model/types'
import { MARGIN_PRESETS } from '../model/pageSetup'
import { setUi } from '../chrome/ui'

export type PageSetupTab = 'Page' | 'Margins' | 'Header/Footer' | 'Sheet'

export function updatePageSetup(patch: Partial<PageSetup>, label = 'Page Setup'): void {
  store.updateLayout(store.active, label, (l) => { l.pageSetup = { ...l.pageSetup, ...patch } })
}

export function openPageSetup(tab: PageSetupTab = 'Page'): void {
  setUi({ backstage: false, dialog: { kind: 'pageSetup', props: { tab } } })
}

export function setMarginPreset(name: keyof typeof MARGIN_PRESETS): void {
  updatePageSetup({ margins: { ...MARGIN_PRESETS[name] } }, 'Margins')
}

/** Scale to Fit > Width / Height: 0 = Automatic, n = n pages. Fit stays off while both are automatic. */
export function setFitPages(axis: 'w' | 'h', pages: number): void {
  const ps = store.active.layout.pageSetup ?? {}
  const w = axis === 'w' ? pages : ps.fitToPage ? ps.fitToWidth ?? 1 : 0
  const h = axis === 'h' ? pages : ps.fitToPage ? ps.fitToHeight ?? 1 : 0
  updatePageSetup({ fitToPage: w !== 0 || h !== 0, fitToWidth: w, fitToHeight: h }, 'Scale to Fit')
}

export function setScale(percent: number): void {
  updatePageSetup({ scale: Math.min(400, Math.max(10, Math.round(percent))), fitToPage: false }, 'Scale')
}

const sorted = (a: number[]): number[] => [...new Set(a)].sort((x, y) => x - y)

/** Breaks > Insert Page Break: a horizontal break above the active row and a vertical one left of the active column. */
export function insertPageBreak(): void {
  const { r, c } = store.active.selection.active
  const ps = store.active.layout.pageSetup ?? {}
  updatePageSetup({
    rowBreaks: r > 0 ? sorted([...(ps.rowBreaks ?? []), r]) : ps.rowBreaks,
    colBreaks: c > 0 ? sorted([...(ps.colBreaks ?? []), c]) : ps.colBreaks
  }, 'Insert Page Break')
}

export function canRemovePageBreak(): boolean {
  const { r, c } = store.active.selection.active
  const ps = store.active.layout.pageSetup
  return !!ps && ((ps.rowBreaks ?? []).includes(r) || (ps.colBreaks ?? []).includes(c))
}

export function removePageBreak(): void {
  const { r, c } = store.active.selection.active
  const ps = store.active.layout.pageSetup ?? {}
  updatePageSetup({ rowBreaks: (ps.rowBreaks ?? []).filter((x) => x !== r), colBreaks: (ps.colBreaks ?? []).filter((x) => x !== c) }, 'Remove Page Break')
}

export function resetPageBreaks(): void {
  updatePageSetup({ rowBreaks: [], colBreaks: [] }, 'Reset All Page Breaks')
}

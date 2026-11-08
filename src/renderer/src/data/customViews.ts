// View > Custom Views: named snapshots of the active sheet's view. Kept for the session only (not written to the file).
import { setUi, getUi } from '../chrome/ui'
import { store } from '../model/store'
import { Selection } from '../model/types'

export interface CustomView {
  name: string
  sheetUid: number
  selection: Selection
  zoom: number
  frozen: { rows: number; cols: number }
  showGridlines: boolean
  showHeadings: boolean
  viewMode: 'normal' | 'pageBreak'
}

let views: CustomView[] = []
export const listViews = (): CustomView[] => views

/** Same name (case-insensitive) replaces, as Excel's Add does after confirmation. */
export function upsertView(list: CustomView[], v: CustomView): CustomView[] {
  const i = list.findIndex((x) => x.name.toLowerCase() === v.name.toLowerCase())
  return i < 0 ? [...list, v] : list.map((x, k) => (k === i ? v : x))
}

export function addView(name: string): void {
  const s = store.active
  const l = s.layout
  const sel = s.selection
  views = upsertView(views, {
    name, sheetUid: s.uid, zoom: l.zoom, frozen: { ...l.frozen }, showGridlines: l.showGridlines, showHeadings: l.showHeadings,
    viewMode: getUi().viewMode === 'pageBreak' ? 'pageBreak' : 'normal',
    selection: { ranges: sel.ranges.map((r) => ({ ...r })), active: { ...sel.active } }
  })
}

export function deleteView(name: string): void { views = views.filter((v) => v.name !== name) }

/** Returns false when the view's sheet no longer exists. */
export function showView(v: CustomView): boolean {
  const idx = store.sheets.findIndex((s) => s.uid === v.sheetUid)
  if (idx < 0) return false
  store.setActiveSheet(idx)
  const s = store.active
  store.updateLayout(s, 'Custom View', (l) => {
    l.zoom = v.zoom; l.frozen = { ...v.frozen }; l.showGridlines = v.showGridlines; l.showHeadings = v.showHeadings
  })
  store.setSelection({ ranges: v.selection.ranges.map((r) => ({ ...r })), active: { ...v.selection.active } }, s)
  setUi({ viewMode: v.viewMode })
  return true
}

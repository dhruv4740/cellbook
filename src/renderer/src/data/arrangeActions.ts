// Ribbon-facing Arrange commands: resolve the selected floating objects, run the pure op inside one undo step.
import { store, SheetState } from '../model/store'
import { getUi, setUi, showToast } from '../chrome/ui'
import * as A from './arrange'

/** Ids of the selected floating objects on the active sheet (empty when none). */
export function selectedObjectIds(): string[] {
  const ui = getUi()
  const prim = ui.selectedPicture ?? ui.selectedChart
  if (!prim) return []
  return ui.selectedObjects.includes(prim) ? ui.selectedObjects : [prim]
}

/** Selects `id` (and its whole group); `additive` toggles it within the current selection (Ctrl/Shift+click). */
export function selectObject(sheet: SheetState, id: string, additive = false): void {
  const all = A.floatsOf(sheet.layout)
  if (!additive && selectedObjectIds().length > 1 && selectedObjectIds().includes(id)) return // dragging a member keeps the multi-selection
  const cur = additive ? selectedObjectIds() : []
  const adding = !cur.includes(id)
  let ids: string[]
  if (adding) ids = A.withGroupMembers(sheet.layout, [...cur, id])
  else {
    const g = all.find((f) => f.id === id)?.obj.group
    ids = cur.filter((x) => x !== id && (!g || all.find((f) => f.id === x)?.obj.group !== g))
  }
  const prim = adding ? id : ids[ids.length - 1]
  const kind = all.find((f) => f.id === prim)?.kind
  setUi({ selectedObjects: ids, selectedPicture: kind && kind !== 'chart' ? prim : null, selectedChart: kind === 'chart' ? prim : null })
}

/** Selects exactly these ids (rubber band, Selection Pane). */
export function selectObjects(sheet: SheetState, ids: string[]): void {
  if (!ids.length) { setUi({ selectedObjects: [], selectedPicture: null, selectedChart: null }); return }
  selectObject(sheet, ids[0])
  for (const id of ids.slice(1)) selectObject(sheet, id, true)
}

function run(label: string, fn: (l: SheetState['layout']) => void): void {
  const sheet = store.active
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, label, fn)
}

const need = (n: number): string[] | null => {
  const ids = selectedObjectIds()
  if (ids.length >= n) return ids
  showToast(n === 1 ? 'Select an object first.' : `Select at least ${n} objects first.`)
  return null
}

export const arrangeOrder = (mode: 'front' | 'back' | 'forward' | 'backward'): void => {
  const ids = need(1)
  if (ids) run(mode === 'front' ? 'Bring to Front' : mode === 'back' ? 'Send to Back' : mode === 'forward' ? 'Bring Forward' : 'Send Backward', (l) => A.reorder(l, ids, mode))
}
export const arrangeAlign = (mode: A.AlignMode): void => { const ids = need(2); if (ids) run('Align Objects', (l) => A.align(l, ids, mode)) }
export const arrangeDistribute = (horizontal: boolean): void => { const ids = need(3); if (ids) run('Distribute Objects', (l) => A.distribute(l, ids, horizontal)) }
export const arrangeGroup = (): void => { const ids = need(2); if (ids) run('Group', (l) => A.group(l, ids, 'grp_' + Date.now().toString(36))) }
export const arrangeUngroup = (): void => { const ids = need(1); if (ids) run('Ungroup', (l) => A.ungroup(l, ids)) }
export const arrangeTransform = (op: { rotate: number } | { flip: 'h' | 'v' }): void => { const ids = need(1); if (ids) run('rotate' in op ? 'Rotate' : 'Flip', (l) => A.transform(l, ids, op)) }
export const setObjectsHidden = (ids: string[], hidden: boolean): void => run(hidden ? 'Hide Object' : 'Show Object', (l) => A.setHidden(l, ids, hidden))
export const setObjectOrder = (order: string[]): void => run('Reorder Objects', (l) => A.setOrder(l, order))

/**
 * A drag that only moved (same size) a member of a multi-selection / group moves all of them by the same delta.
 * Returns true when it handled the commit.
 */
export function commitMoveOfSelection(sheet: SheetState, id: string, anchor: { r: number; c: number; dx: number; dy: number }, width: number, height: number): boolean {
  const ids = selectedObjectIds()
  if (ids.length < 2 || !ids.includes(id)) return false
  const f = A.floatsOf(sheet.layout).find((x) => x.id === id)
  if (!f || Math.abs(f.obj.width - width) > 0.5 || Math.abs(f.obj.height - height) > 0.5) return false
  const ax = A.axesOf(sheet.layout)
  const a = A.absRect(f.obj, ax), b = A.absRect({ ...f.obj, anchor }, ax)
  run('Move Objects', (l) => A.shiftBy(l, ids, b.x - a.x, b.y - a.y))
  return true
}

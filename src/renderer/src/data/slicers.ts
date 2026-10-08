// Insert > Slicer / Timeline: floating filter panels that drive the sheet's AutoFilter on one column.
import { store, SheetState } from '../model/store'
import { SlicerDef } from '../model/types'
import { BLANK_TOKEN } from '../engine/filter'
import { setUi, showToast } from '../chrome/ui'

export type TimelinePeriod = NonNullable<SlicerDef['period']>

export const slicersOf = (sheet: SheetState): SlicerDef[] => sheet.layout.slicers ?? []

const serialOfDate = (y: number, m: number, d: number): number => Math.round(Date.UTC(y, m, d) / 86400000) + 25569
const dateOfSerial = (n: number): Date => new Date(Math.round((n - 25569) * 86400000))
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

/** Distinct displayed values of the slicer's column across the whole filter range (including rows the filter hides). */
export function slicerItems(sheet: SheetState, def: SlicerDef): string[] {
  const af = sheet.layout.autoFilter
  if (!af) return []
  const seen = new Set<string>()
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) { const t = store.getDisplay(sheet, r, def.col).text; seen.add(t === '' ? BLANK_TOKEN : t) }
  const isNum = [...seen].every((t) => t === BLANK_TOKEN || (t.trim() !== '' && !isNaN(Number(t))))
  return [...seen].sort((a, b) => (a === BLANK_TOKEN ? 1 : b === BLANK_TOKEN ? -1 : isNum ? Number(a) - Number(b) : a.localeCompare(b)))
}

/** Serial of the first and last row's date value, or null when the column holds no dates. */
export function timelineSpan(sheet: SheetState, def: SlicerDef): [number, number] | null {
  const af = sheet.layout.autoFilter
  if (!af) return null
  let lo = Infinity, hi = -Infinity
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) { const v = store.getValue(sheet, r, def.col); if (typeof v === 'number') { lo = Math.min(lo, Math.floor(v)); hi = Math.max(hi, Math.floor(v)) } }
  return lo === Infinity ? null : [lo, hi]
}

/** The period buttons of a timeline at its current level: label plus the inclusive serial range each covers. */
export function timelinePeriods(span: [number, number], period: TimelinePeriod): { label: string; lo: number; hi: number }[] {
  const out: { label: string; lo: number; hi: number }[] = []
  const a = dateOfSerial(span[0]), b = dateOfSerial(span[1])
  if (period === 'days') { for (let s = span[0]; s <= span[1] && out.length < 400; s++) out.push({ label: String(dateOfSerial(s).getUTCDate()), lo: s, hi: s }); return out }
  const step = period === 'years' ? 12 : period === 'quarters' ? 3 : 1
  let m = a.getUTCFullYear() * 12 + (period === 'quarters' ? Math.floor(a.getUTCMonth() / 3) * 3 : period === 'years' ? 0 : a.getUTCMonth())
  const last = b.getUTCFullYear() * 12 + b.getUTCMonth()
  for (; m <= last; m += step) {
    const y = Math.floor(m / 12), mo = m % 12
    out.push({ label: period === 'years' ? String(y) : period === 'quarters' ? `Q${mo / 3 + 1} ${y}` : `${MONTHS[mo]} ${y}`, lo: serialOfDate(y, mo, 1), hi: serialOfDate(y, mo + step, 1) - 1 })
  }
  return out
}

/** Writes the slicer / timeline selection into the AutoFilter column (all items selected = no criteria). */
export function applySlicer(sheet: SheetState, def: SlicerDef): void {
  const af = sheet.layout.autoFilter
  if (!af) return
  let selected: string[] | null = null
  if (def.kind === 'timeline') {
    if (def.range) {
      const [lo, hi] = def.range
      const set = new Set<string>()
      for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
        const v = store.getValue(sheet, r, def.col)
        const n = typeof v === 'number' ? Math.floor(v) : NaN
        if (n >= lo && n <= hi) { const t = store.getDisplay(sheet, r, def.col).text; set.add(t === '' ? BLANK_TOKEN : t) }
      }
      selected = [...set]
    }
  } else if (def.selected) selected = def.selected
  store.setColumnFilter(sheet, def.col, selected ? { selected } : null)
}

export function updateSlicer(sheet: SheetState, id: string, patch: Partial<SlicerDef>, label = 'Slicer'): void {
  store.updateLayout(sheet, label, (l) => { l.slicers = (l.slicers ?? []).map((s) => (s.id === id ? { ...s, ...patch } : s)) })
}

export function deleteSlicer(sheet: SheetState, id: string): void {
  const def = slicersOf(sheet).find((s) => s.id === id)
  store.updateLayout(sheet, 'Delete Slicer', (l) => { l.slicers = (l.slicers ?? []).filter((s) => s.id !== id) })
  if (def) store.setColumnFilter(sheet, def.col, null) // removing a slicer drops its filter, like Excel
}

/** Select an item: plain click shows only it; Ctrl toggles it in the set; clicking the only selected item shows all. */
export function pickSlicerItem(sheet: SheetState, def: SlicerDef, item: string, additive: boolean): void {
  const all = slicerItems(sheet, def)
  const cur = def.selected ?? all
  let next: string[]
  if (additive) next = cur.includes(item) ? cur.filter((x) => x !== item) : [...cur, item]
  else next = cur.length === 1 && cur[0] === item ? all : [item]
  const selected = next.length === 0 || next.length === all.length ? null : next
  const nd = { ...def, selected }
  updateSlicer(sheet, def.id, { selected })
  applySlicer(sheet, nd)
}

/** Timeline: a click selects one period; Shift extends the range from the previous pick. */
export function pickTimelinePeriod(sheet: SheetState, def: SlicerDef, p: { lo: number; hi: number }, extend: boolean): void {
  let range: [number, number] | null = [p.lo, p.hi]
  if (extend && def.range) range = [Math.min(def.range[0], p.lo), Math.max(def.range[1], p.hi)]
  else if (def.range && def.range[0] === p.lo && def.range[1] === p.hi) range = null
  updateSlicer(sheet, def.id, { range })
  applySlicer(sheet, { ...def, range })
}

export function clearSlicer(sheet: SheetState, def: SlicerDef): void {
  updateSlicer(sheet, def.id, { selected: null, range: null }, 'Clear Slicer Filter')
  applySlicer(sheet, { ...def, selected: null, range: null })
}

/** Insert > Slicer / Timeline for the active cell's column. Adds an AutoFilter over the current region first when there is none. */
export function insertSlicer(kind: 'slicer' | 'timeline'): void {
  const sheet = store.active
  if (!store.allows(sheet, 'objects')) return
  if (!sheet.layout.autoFilter) store.toggleAutoFilter(sheet)
  const af = store.active.layout.autoFilter
  if (!af) return
  const { r, c } = sheet.selection.active
  const col = c >= af.range.c1 && c <= af.range.c2 ? c : af.range.c1
  const probe: SlicerDef = { id: '', name: '', kind, col, header: '', anchor: { r: 0, c: 0, dx: 0, dy: 0 }, width: 0, height: 0 }
  if (kind === 'timeline' && !timelineSpan(store.active, probe)) { showToast('Timeline needs a column of dates. Select a cell in a date column first.'); return }
  const header = store.getDisplay(sheet, af.range.r1, col).text || `Column ${col + 1}`
  const n = slicersOf(sheet).filter((s) => s.kind === kind).length + 1
  const def: SlicerDef = {
    id: 'slc_' + Date.now().toString(36) + n, name: `${header}${n > 1 ? ' ' + n : ''}`, kind, col, header,
    anchor: { r: Math.max(0, r), c: Math.min(af.range.c2 + 2, 16000), dx: 24 * slicersOf(sheet).length, dy: 24 * slicersOf(sheet).length }, // cascade so a new panel never hides the last
    width: kind === 'timeline' ? 360 : 160, height: kind === 'timeline' ? 120 : 220, period: kind === 'timeline' ? 'months' : undefined
  }
  store.updateLayout(sheet, kind === 'timeline' ? 'Insert Timeline' : 'Insert Slicer', (l) => { l.slicers = [...(l.slicers ?? []), def] })
  setUi({ selectedPicture: null })
}

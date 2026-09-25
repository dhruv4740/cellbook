// Moves everything in a sheet's layout that is addressed by cell when rows/columns are inserted
// (delta > 0) or deleted (delta < 0) at index `at`, following Excel: ranges that span the insertion
// point grow, ranges cut by a deletion shrink, and anything wholly inside a deleted band goes away.
import type { Range, SheetLayout } from './types'

type Axis = 'r' | 'c'

/** New index for a single row/column, or null when it was deleted. */
export function shiftIndex(i: number, at: number, delta: number): number | null {
  if (i < at) return i
  if (delta > 0) return i + delta
  if (i < at - delta) return null
  return i + delta
}

/** Shifted copy of a range along one axis, or null when the deletion removed all of it. */
export function shiftRange(rg: Range, axis: Axis, at: number, delta: number): Range | null {
  const [a, b] = axis === 'r' ? (['r1', 'r2'] as const) : (['c1', 'c2'] as const)
  const lo = Math.min(rg[a], rg[b]), hi = Math.max(rg[a], rg[b])
  let nlo: number, nhi: number
  if (delta > 0) {
    nlo = lo >= at ? lo + delta : lo
    nhi = hi >= at ? hi + delta : hi
  } else {
    const end = at - delta // first index after the deleted band
    if (lo >= at && hi < end) return null
    nlo = lo < at ? lo : lo < end ? at : lo + delta
    nhi = hi < at ? hi : hi < end ? at - 1 : hi + delta
  }
  return { ...rg, [a]: nlo, [b]: nhi }
}

function shiftKeyed<T>(map: Record<string, T> | undefined, axis: Axis, at: number, delta: number): Record<string, T> {
  const out: Record<string, T> = {}
  for (const [k, v] of Object.entries(map ?? {})) {
    const [r, c] = k.split(',').map(Number)
    const i = shiftIndex(axis === 'r' ? r : c, at, delta)
    if (i === null) continue
    out[axis === 'r' ? `${i},${c}` : `${r},${i}`] = v
  }
  return out
}

function shiftRanges(ranges: Range[], axis: Axis, at: number, delta: number): Range[] {
  return ranges.map((rg) => shiftRange(rg, axis, at, delta)).filter((x): x is Range => x !== null)
}

export function shiftLayoutRefs(l: SheetLayout, axis: Axis, at: number, delta: number): void {
  l.notes = shiftKeyed(l.notes, axis, at, delta)
  l.comments = shiftKeyed(l.comments, axis, at, delta)
  const ext = l as SheetLayout & { hyperlinks?: Record<string, unknown> }
  if (ext.hyperlinks) ext.hyperlinks = shiftKeyed(ext.hyperlinks, axis, at, delta)
  if (l.conditionalFormats) {
    l.conditionalFormats = l.conditionalFormats
      .map((cf) => ({ ...cf, ranges: shiftRanges(cf.ranges, axis, at, delta) }))
      .filter((cf) => cf.ranges.length)
  }
  if (l.dataValidations) {
    l.dataValidations = l.dataValidations
      .map((dv) => ({ ...dv, ranges: shiftRanges(dv.ranges, axis, at, delta) }))
      .filter((dv) => dv.ranges.length)
  }
  if (l.autoFilter) {
    const header = axis === 'r' ? l.autoFilter.range.r1 : l.autoFilter.range.c1
    const rg = shiftRange(l.autoFilter.range, axis, at, delta)
    if (!rg || shiftIndex(header, at, delta) === null) {
      l.autoFilter = undefined // header row/column deleted: Excel removes the filter
    } else {
      const af = { ...l.autoFilter, range: rg }
      if (axis === 'r') {
        af.filteredOutRows = af.filteredOutRows.map((r) => shiftIndex(r, at, delta)).filter((r): r is number => r !== null)
      } else {
        const cols: typeof af.columns = {}
        for (const [c, crit] of Object.entries(af.columns)) {
          const n = shiftIndex(+c, at, delta)
          if (n !== null) cols[n] = crit
        }
        af.columns = cols
      }
      l.autoFilter = af
    }
  }
  if (l.charts) {
    l.charts = l.charts.map((ch) => {
      const dataRange = shiftRange(ch.dataRange, axis, at, delta) ?? ch.dataRange
      const i = shiftIndex(axis === 'r' ? ch.anchor.r : ch.anchor.c, at, delta) ?? at
      return { ...ch, dataRange, anchor: axis === 'r' ? { ...ch.anchor, r: i } : { ...ch.anchor, c: i } }
    })
  }
}

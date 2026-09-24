// Pure AutoFilter logic: which rows a set of per-column checklist criteria hides. Kept free of
// the store/HyperFormula so it's cheap to unit test; Grid/store call this with a getDisplay reader.
import { Range, FilterCriteria } from '../model/types'

/** Sentinel standing in for an empty cell in a column's checked-values list. */
export const BLANK_TOKEN = '\u0000blank\u0000'

/** Every distinct display value in `range`'s column `col`, below the header row. */
export function distinctColumnValues(range: Range, col: number, getDisplay: (r: number, c: number) => string): string[] {
  const seen = new Set<string>()
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    const t = getDisplay(r, col)
    seen.add(t === '' ? BLANK_TOKEN : t)
  }
  return [...seen]
}

/** Rows (below the header) that fail at least one column's criteria, plus how many pass. */
export function computeFilteredRows(
  range: Range,
  columns: Record<number, FilterCriteria>,
  getDisplay: (r: number, c: number) => string
): { hidden: number[]; visible: number; total: number } {
  const hidden: number[] = []
  let visible = 0
  const total = Math.max(0, range.r2 - range.r1)
  const cols = Object.entries(columns).map(([c, crit]) => [Number(c), new Set(crit.selected)] as const)
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    let show = true
    for (const [c, selected] of cols) {
      const t = getDisplay(r, c)
      const token = t === '' ? BLANK_TOKEN : t
      if (!selected.has(token)) { show = false; break }
    }
    if (show) { visible++ } else { hidden.push(r) }
  }
  return { hidden, visible, total }
}

export type NumberComparator =
  | { op: 'eq' | 'ne' | 'gt' | 'gte' | 'lt' | 'lte'; value: number }
  | { op: 'between'; value: number; value2: number }
  | { op: 'top' | 'bottom'; count: number; percent?: boolean }
  | { op: 'aboveAverage' | 'belowAverage' }

export type TextComparator =
  | { op: 'eq' | 'ne' | 'contains' | 'notContains' | 'beginsWith' | 'endsWith'; value: string }

/** Builds the `selected` list for "select all" / "select none" / toggling one value. */
export function toggleValue(selected: string[], value: string, allValues: string[]): string[] {
  const s = new Set(selected)
  if (s.has(value)) s.delete(value)
  else s.add(value)
  return allValues.filter((v) => s.has(v))
}

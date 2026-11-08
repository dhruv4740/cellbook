// Data > Group / Ungroup outlines for rows and columns. Pure functions over SheetLayout so they can be
// undone through store.updateLayout and shared by the ribbon, the gutter and the file reader/writer.
//
// A line's level is its group depth (0 = not grouped; Excel's file attribute outlineLevel). A "run" is a
// maximal stretch of adjacent lines whose level is >= k; each run has a bracket in the gutter and a [-]/[+]
// button on its summary line (the line after the run, or before it when summaries are above/left).
import { SheetLayout } from './types'

export type Axis = 'row' | 'col'
export const MAX_OUTLINE_LEVEL = 7

export interface OutlineRun {
  level: number
  start: number
  end: number
  /** Line carrying the [-]/[+] button; -1 when there is none (summary above the first line). */
  summary: number
  /** True when every line at exactly this level is hidden. */
  collapsed: boolean
}

function maps(l: SheetLayout, axis: Axis) {
  return axis === 'row'
    ? { levels: l.rowLevels ??= {}, hidden: l.hiddenRows, collapsed: l.collapsedRows ??= {}, after: l.outline?.summaryBelow !== false }
    : { levels: l.colLevels ??= {}, hidden: l.hiddenCols, collapsed: l.collapsedCols ??= {}, after: l.outline?.summaryRight !== false }
}

export function lineLevel(l: SheetLayout, axis: Axis, i: number): number {
  return (axis === 'row' ? l.rowLevels : l.colLevels)?.[i] ?? 0
}

export function maxLevel(l: SheetLayout, axis: Axis): number {
  const m = axis === 'row' ? l.rowLevels : l.colLevels
  let mx = 0
  if (m) for (const k in m) if (m[k] > mx) mx = m[k]
  return mx
}

/** Every bracket of one axis, outermost level first. */
export function outlineRuns(l: SheetLayout, axis: Axis): OutlineRun[] {
  const levels = axis === 'row' ? l.rowLevels : l.colLevels
  if (!levels) return []
  const hidden = axis === 'row' ? l.hiddenRows : l.hiddenCols
  const after = axis === 'row' ? l.outline?.summaryBelow !== false : l.outline?.summaryRight !== false
  const idx = Object.keys(levels).map(Number).filter((i) => levels[i] > 0).sort((a, b) => a - b)
  const out: OutlineRun[] = []
  let mx = 0
  for (const i of idx) if (levels[i] > mx) mx = levels[i]
  for (let k = 1; k <= mx; k++) {
    let start = -1, prev = -2
    const flush = (end: number): void => {
      if (start < 0) return
      let all = true
      for (let i = start; i <= end && all; i++) if ((levels[i] ?? 0) === k && !hidden[i]) all = false
      out.push({ level: k, start, end, summary: after ? end + 1 : start - 1, collapsed: all })
      start = -1
    }
    for (const i of idx) {
      if (levels[i] < k) continue
      if (i !== prev + 1) { flush(prev); start = i }
      prev = i
    }
    flush(prev)
  }
  return out
}

/** Group: adds one level to each line in [from,to]. Returns false at the maximum depth. */
export function groupLines(l: SheetLayout, axis: Axis, from: number, to: number): boolean {
  const { levels } = maps(l, axis)
  for (let i = from; i <= to; i++) if ((levels[i] ?? 0) >= MAX_OUTLINE_LEVEL) return false
  for (let i = from; i <= to; i++) levels[i] = (levels[i] ?? 0) + 1
  return true
}

/** Ungroup: removes one level from each line in [from,to]. Hidden lines stay hidden, like Excel. */
export function ungroupLines(l: SheetLayout, axis: Axis, from: number, to: number): void {
  const { levels } = maps(l, axis)
  for (let i = from; i <= to; i++) {
    const v = levels[i] ?? 0
    if (v > 1) levels[i] = v - 1
    else delete levels[i]
  }
  pruneCollapsed(l, axis)
}

export function clearOutline(l: SheetLayout, axis: Axis): void {
  if (axis === 'row') { delete l.rowLevels; delete l.collapsedRows } else { delete l.colLevels; delete l.collapsedCols }
}

/** Forgets collapsed markers whose summary line no longer ends a group. */
export function pruneCollapsed(l: SheetLayout, axis: Axis): void {
  const { collapsed } = maps(l, axis)
  const summaries = new Set(outlineRuns(l, axis).map((r) => r.summary))
  for (const k of Object.keys(collapsed)) if (!summaries.has(+k)) delete collapsed[+k]
  if (axis === 'row' && !Object.keys(collapsed).length) delete l.collapsedRows
  if (axis === 'col' && !Object.keys(collapsed).length) delete l.collapsedCols
}

/** Collapse ([-] clicked) or expand ([+] clicked) one bracket. */
export function setRunCollapsed(l: SheetLayout, axis: Axis, run: OutlineRun, collapse: boolean): void {
  const { levels, hidden, collapsed } = maps(l, axis)
  if (collapse) {
    for (let i = run.start; i <= run.end; i++) hidden[i] = true
    if (run.summary >= 0) collapsed[run.summary] = true
    return
  }
  if (run.summary >= 0) delete collapsed[run.summary]
  // Lines at this level show; deeper lines show unless one of their own enclosing groups is still collapsed.
  const inner = outlineRuns(l, axis).filter((r) => r.level > run.level && r.start >= run.start && r.end <= run.end)
  for (let i = run.start; i <= run.end; i++) {
    const lv = levels[i] ?? 0
    const blocked = lv > run.level && inner.some((r) => i >= r.start && i <= r.end && r.level <= lv && r.summary >= 0 && collapsed[r.summary])
    if (!blocked) delete hidden[i]
  }
  if (axis === 'row' && !Object.keys(collapsed).length) delete l.collapsedRows
  if (axis === 'col' && !Object.keys(collapsed).length) delete l.collapsedCols
}

/** The 1..n level buttons: n = shown depth + 1 (button 1 shows only ungrouped lines). */
export function showLevels(l: SheetLayout, axis: Axis, n: number): void {
  const { levels, hidden, collapsed } = maps(l, axis)
  for (const k of Object.keys(levels)) {
    if (levels[+k] >= n) hidden[+k] = true
    else delete hidden[+k]
  }
  for (const k of Object.keys(collapsed)) delete collapsed[+k]
  for (const r of outlineRuns(l, axis)) if (r.level >= n && r.summary >= 0) collapsed[r.summary] = true
  if (axis === 'row' && !Object.keys(collapsed).length) delete l.collapsedRows
  if (axis === 'col' && !Object.keys(collapsed).length) delete l.collapsedCols
}

/** The run a Show/Hide Detail command acts on for line i: the one whose summary is i, else the innermost that contains i. */
export function runFor(l: SheetLayout, axis: Axis, i: number): OutlineRun | undefined {
  const runs = outlineRuns(l, axis)
  const bySummary = runs.filter((r) => r.summary === i).sort((a, b) => b.level - a.level)[0]
  if (bySummary) return bySummary
  return runs.filter((r) => i >= r.start && i <= r.end).sort((a, b) => b.level - a.level)[0]
}

// Data > Outline commands (Group, Ungroup, Show/Hide Detail, Auto Outline, Clear Outline) and the gutter buttons.
import { store } from '../model/store'
import { setUi, showToast } from '../chrome/ui'
import { MAX_COLS, MAX_ROWS, normRange, Range } from '../model/types'
import {
  Axis, MAX_OUTLINE_LEVEL, OutlineRun, clearOutline, groupLines, runFor, setRunCollapsed, showLevels, ungroupLines
} from '../model/outline'

export type OutlineMode = 'group' | 'ungroup' | 'show' | 'hide'

function activeRange(): Range {
  const s = store.active.selection
  return normRange(s.ranges[s.ranges.length - 1])
}

/** 'row' / 'col' when the selection is whole rows / whole columns; null for a cell range (Excel then asks). */
export function selectionAxis(): Axis | null {
  const rg = activeRange()
  if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) return 'row'
  if (rg.r1 === 0 && rg.r2 === MAX_ROWS - 1) return 'col'
  return null
}

function span(axis: Axis): [number, number] {
  const rg = activeRange()
  return axis === 'row' ? [rg.r1, rg.r2] : [rg.c1, rg.c2]
}

/** Runs the command for an axis. Used directly for whole rows/columns and by the Rows/Columns dialog. */
export function applyOutline(mode: OutlineMode, axis: Axis): void {
  const sheet = store.active
  if (store.refuseIfProtected(sheet)) return
  const [from, to] = span(axis)
  if (mode === 'group') {
    let ok = true
    store.updateLayout(sheet, 'Group', (l) => { ok = groupLines(l, axis, from, to) })
    if (!ok) showToast(`Outlines can have at most ${MAX_OUTLINE_LEVEL + 1} levels.`)
  } else if (mode === 'ungroup') {
    store.updateLayout(sheet, 'Ungroup', (l) => ungroupLines(l, axis, from, to))
  } else {
    let run: OutlineRun | undefined
    for (let i = from; i <= to && !run; i++) run = runFor(sheet.layout, axis, i)
    if (!run) return
    const r = run
    store.updateLayout(sheet, mode === 'show' ? 'Show Detail' : 'Hide Detail', (l) => setRunCollapsed(l, axis, r, mode === 'hide'))
  }
}

/** Ribbon / keyboard entry: acts on whole rows/columns directly, otherwise opens Excel's Rows/Columns dialog. */
export function outlineCommand(mode: OutlineMode): void {
  const axis = selectionAxis()
  if (axis) applyOutline(mode, axis)
  else setUi({ dialog: { kind: 'outlineAxis', props: { mode } } })
}

/** Gutter: the numbered level buttons. `n` is the button number (1 = show only ungrouped lines). */
export function showOutlineLevel(axis: Axis, n: number): void {
  store.updateLayout(store.active, 'Show Levels', (l) => showLevels(l, axis, n))
}

export function toggleRun(axis: Axis, run: OutlineRun): void {
  store.updateLayout(store.active, run.collapsed ? 'Show Detail' : 'Hide Detail', (l) => setRunCollapsed(l, axis, run, !run.collapsed))
}

export function clearAllOutlines(): void {
  const sheet = store.active
  if (store.refuseIfProtected(sheet)) return
  store.updateLayout(sheet, 'Clear Outline', (l) => { clearOutline(l, 'row'); clearOutline(l, 'col') })
}

export function setOutlineSummary(summaryBelow: boolean, summaryRight: boolean): void {
  store.updateLayout(store.active, 'Outline Settings', (l) => { l.outline = { summaryBelow, summaryRight } })
}

/** Data > Group > Auto Outline: groups the detail lines that formulas summarise, using the sheet's summary direction. */
export function autoOutline(): void {
  const sheet = store.active
  if (store.refuseIfProtected(sheet)) return
  const below = sheet.layout.outline?.summaryBelow !== false
  const right = sheet.layout.outline?.summaryRight !== false
  const rowGroups = new Map<string, [number, number]>()
  const colGroups = new Map<string, [number, number]>()
  const used = store.usedRange(sheet)
  if (!used) return
  const refRe = /(?<![A-Za-z0-9_!'])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g
  const colNum = (s: string): number => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1 }
  for (let r = 0; r <= used.r2; r++) {
    for (let c = 0; c <= used.c2; c++) {
      const raw = store.getRaw(sheet, r, c)
      if (!raw.startsWith('=') || raw.includes('!')) continue
      let m: RegExpExecArray | null
      refRe.lastIndex = 0
      while ((m = refRe.exec(raw))) {
        const c1 = colNum(m[1]), r1 = +m[2] - 1
        const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1
        // Detail rows sit directly above (or below) the summary row in the same column(s).
        if (r2 > r1 && c1 === c2 && ((below && r2 === r - 1 && c1 === c) || (!below && r1 === r + 1 && c1 === c))) rowGroups.set(r1 + ',' + r2, [r1, r2])
        if (c2 > c1 && r1 === r2 && ((right && c2 === c - 1 && r1 === r) || (!right && c1 === c + 1 && r1 === r))) colGroups.set(c1 + ',' + c2, [c1, c2])
      }
    }
  }
  if (!rowGroups.size && !colGroups.size) { setUi({ dialog: { kind: 'message', props: { text: 'Cannot create an outline.' } } }); return }
  store.updateLayout(sheet, 'Auto Outline', (l) => {
    // Outer (longer) groups first so nesting depth follows containment.
    const byLen = (a: [number, number], b: [number, number]): number => (b[1] - b[0]) - (a[1] - a[0])
    for (const [a, b] of [...rowGroups.values()].sort(byLen)) groupLines(l, 'row', a, b)
    for (const [a, b] of [...colGroups.values()].sort(byLen)) groupLines(l, 'col', a, b)
  })
}

// Formulas > Trace Precedents / Trace Dependents: blue arrows between a cell and the cells it reads or feeds.
// Each further click on the same cell goes one level deeper, like Excel.
import { store } from '../model/store'
import { getUi, setUi } from '../chrome/ui'
import type { Range } from '../model/types'

export interface TraceArrow { from: Range; to: Range }
export interface TraceState { sheetUid: number; arrows: TraceArrow[]; levels: Record<string, number> }

type Addr = { sheet: number; row: number; col: number }
type HfRange = { start: Addr; end: Addr }
const isRange = (x: Addr | HfRange): x is HfRange => 'start' in x
const toRange = (x: Addr | HfRange): Range => (isRange(x) ? { r1: x.start.row, c1: x.start.col, r2: x.end.row, c2: x.end.col } : { r1: x.row, c1: x.col, r2: x.row, c2: x.col })
const sheetOf = (x: Addr | HfRange): number => (isRange(x) ? x.start.sheet : x.sheet)
const cellRange = (r: number, c: number): Range => ({ r1: r, c1: c, r2: r, c2: c })
const key = (a: TraceArrow): string => `${a.from.r1},${a.from.c1},${a.from.r2},${a.from.c2}>${a.to.r1},${a.to.c1}`

function related(addr: Addr | HfRange, dir: 'precedents' | 'dependents'): (Addr | HfRange)[] {
  try { return dir === 'precedents' ? store.hf.getCellPrecedents(addr) : store.hf.getCellDependents(addr) } catch { return [] }
}

/** Arrows from the cell down `depth` levels (breadth-first, this sheet only). */
export function traceArrows(r: number, c: number, dir: 'precedents' | 'dependents', depth: number): TraceArrow[] {
  const sheet = store.active
  const sid = store.hfId(sheet)
  const out: TraceArrow[] = []
  const seen = new Set<string>()
  let frontier: [number, number][] = [[r, c]]
  for (let level = 0; level < depth && frontier.length; level++) {
    const next: [number, number][] = []
    for (const [fr, fc] of frontier) {
      // A range vertex in the dependents list stands for the formulas that read that range: use those instead.
      let rels = related({ sheet: sid, row: fr, col: fc }, dir)
      if (dir === 'dependents') rels = rels.flatMap((x) => (isRange(x) ? related(x, dir) : [x]))
      for (const rel of rels) {
        if (sheetOf(rel) !== sid) continue
        const rg = toRange(rel)
        const arrow: TraceArrow = dir === 'precedents' ? { from: rg, to: cellRange(fr, fc) } : { from: cellRange(fr, fc), to: rg }
        const k = key(arrow)
        if (seen.has(k)) continue
        seen.add(k)
        out.push(arrow)
        const n = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1)
        if (n <= 100) for (let rr = rg.r1; rr <= rg.r2; rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) next.push([rr, cc])
      }
    }
    frontier = next
  }
  return out
}

export function traceStep(dir: 'precedents' | 'dependents'): void {
  const sheet = store.active
  const { r, c } = sheet.selection.active
  const cur = getUi().traceArrows
  const state: TraceState = cur && cur.sheetUid === sheet.uid ? cur : { sheetUid: sheet.uid, arrows: [], levels: {} }
  const lk = `${dir}:${r},${c}`
  const depth = (state.levels[lk] ?? 0) + 1
  const arrows = traceArrows(r, c, dir, depth)
  const had = new Set(state.arrows.map(key))
  const fresh = arrows.filter((a) => !had.has(key(a)))
  if (!fresh.length) {
    if (!state.arrows.length) {
      const msg = dir === 'precedents'
        ? 'The Trace Precedents command requires that the active cell contain a formula which refers to other cells.'
        : 'The Trace Dependents command requires that the active cell be referred to by a formula in the workbook.'
      setUi({ dialog: { kind: 'message', props: { text: msg } } })
    }
    return
  }
  setUi({ traceArrows: { sheetUid: sheet.uid, arrows: [...state.arrows, ...fresh], levels: { ...state.levels, [lk]: depth } } })
}

export function removeArrows(): void {
  setUi({ traceArrows: null })
}

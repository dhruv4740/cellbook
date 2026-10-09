// Performance patch for HyperFormula's dependency graph.
//
// When many formulas share one range (VLOOKUP(x, $A$2:$B$50000, ...) down a column), HyperFormula re-adds an edge from
// every cell of that range to the shared range vertex for EACH formula: O(rows) per formula, O(rows^2) to load or fill
// a lookup column (5,000 lookups took 5 s, 50,000 would take minutes). The vertex is wired completely when it is first
// created, so for a range that already exists we only need the edge from the range to the new formula.
import { HyperFormula } from 'hyperformula'
import { bumpAggEpoch } from './fastAgg'

let patched = false
export function patchHyperFormulaGraph(): void {
  if (patched) return
  patched = true
  try {
    const probe: any = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
    const proto = Object.getPrototypeOf(probe.dependencyGraph)
    const orig = proto.processCellDependencies
    if (typeof orig !== 'function') return
    proto.processCellDependencies = function (this: any, deps: any[], endVertex: any): void {
      let rest: any[] | null = null
      for (let i = 0; i < deps.length; i++) {
        const dep = deps[i]
        if (dep && dep.start && dep.end && typeof dep.isFinite === 'function' && dep.isFinite()) {
          const rv = this.getRange(dep.start, dep.end)
          if (rv !== undefined && rv.idInGraph !== undefined && endVertex.idInGraph !== undefined && this.graph.hasNode(rv)) {
            this.graph.addEdge(rv.idInGraph, endVertex.idInGraph)
            continue
          }
        }
        if (rest === null) rest = []
        rest.push(dep)
      }
      if (rest !== null && rest.length) orig.call(this, rest, endVertex)
    }
    // Recalculation seeds every changed cell. A plain value (or empty) cell that nothing depends on has nothing to recompute,
    // so leave it out of the topological sort: writing a million constants (sort, paste, dedupe) skipped a million DFS runs.
    const gp = Object.getPrototypeOf(probe.dependencyGraph.graph)
    const origTop = gp.getTopSortedWithSccSubgraphFrom
    if (typeof origTop === 'function') {
      gp.getTopSortedWithSccSubgraphFrom = function (this: any, nodes: any[], op: unknown, onCycle: unknown): unknown {
        const keep = nodes.filter((n) => {
          if (n.idInGraph === undefined) return false
          const e = this.edgesSparseArray[n.idInGraph]
          return (e !== undefined && e.length > 0) || typeof n.isComputed === 'function' || typeof n.clearCache === 'function'
        })
        return origTop.call(this, keep, op, onCycle)
      }
    }
    // HyperFormula records every changed cell so it can report them to 'valuesUpdated' listeners and as the return value of
    // setCellContents. Cellbook reads values straight from the engine and uses neither, and the bookkeeping cost a second
    // per million cells written.
    const cc = probe._evaluator?.partialRun?.([])
    const ccp = cc ? Object.getPrototypeOf(cc) : null
    if (ccp && typeof ccp.add === 'function') ccp.add = function (): void { /* untracked */ }
    // Parsing raw text (number? date? percentage? boolean?) dominates bulk writes such as sorting; the same strings recur
    // constantly (station codes, statuses), so remember the parse result per string.
    const cp = Object.getPrototypeOf(probe._cellContentParser)
    const origParse = cp.parse
    const caches = new WeakMap<object, Map<string, unknown>>()
    cp.parse = function (this: object, content: unknown): unknown {
      if (typeof content !== 'string' || content.length > 64 || content.charCodeAt(0) === 61) return origParse.call(this, content)
      let m = caches.get(this)
      if (!m) caches.set(this, (m = new Map()))
      const hit = m.get(content)
      if (hit !== undefined) return hit
      const res = origParse.call(this, content)
      if (m.size > 200_000) m.clear()
      m.set(content, res)
      return res
    }
    probe.destroy()
  } catch { /* unpatched HyperFormula is only slower */ }
}

const MUTATORS = [
  'setCellContents', 'setSheetContent', 'addSheet', 'removeSheet', 'renameSheet', 'clearSheet', 'addRows', 'removeRows', 'addColumns', 'removeColumns',
  'moveCells', 'moveRows', 'moveColumns', 'paste', 'cut', 'rebuildAndRecalculate', 'batch', 'undo', 'redo', 'resumeEvaluation'
]
/** Makes every engine mutation start a new cache epoch for the fast aggregate functions (see fastAgg.ts). */
export function instrumentHf(hf: HyperFormula): HyperFormula {
  const h = hf as any
  for (const m of MUTATORS) {
    const orig = h[m]
    if (typeof orig !== 'function') continue
    h[m] = function (...args: unknown[]) {
      bumpAggEpoch()
      try { return orig.apply(hf, args) } finally { bumpAggEpoch() }
    }
  }
  return hf
}

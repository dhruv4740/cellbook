// Data > Consolidate: combines several ranges into one summary, by position or by row/column labels.
export type ConsolidateFn = 'Sum' | 'Count' | 'Average' | 'Max' | 'Min' | 'Product' | 'Count Numbers' | 'StdDev' | 'StdDevp' | 'Var' | 'Varp'
export const CONSOLIDATE_FUNCTIONS: ConsolidateFn[] = ['Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'Count Numbers', 'StdDev', 'StdDevp', 'Var', 'Varp']

export type Cell = string | number | boolean | null
export interface Source { data: Cell[][] }

function combine(fn: ConsolidateFn, xs: Cell[]): number | null {
  const nums = xs.filter((x): x is number => typeof x === 'number')
  const n = nums.length
  const sum = nums.reduce((a, b) => a + b, 0)
  const mean = n ? sum / n : 0
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0)
  switch (fn) {
    case 'Sum': return n ? sum : null
    case 'Count': return xs.filter((x) => x !== null && x !== '').length || null
    case 'Count Numbers': return n || null
    case 'Average': return n ? mean : null
    case 'Max': return n ? Math.max(...nums) : null
    case 'Min': return n ? Math.min(...nums) : null
    case 'Product': return n ? nums.reduce((a, b) => a * b, 1) : null
    case 'StdDev': return n > 1 ? Math.sqrt(ss / (n - 1)) : null
    case 'StdDevp': return n ? Math.sqrt(ss / n) : null
    case 'Var': return n > 1 ? ss / (n - 1) : null
    case 'Varp': return n ? ss / n : null
  }
}

/** By position: cell (i,j) of the result combines cell (i,j) of every source. */
export function consolidateByPosition(sources: Source[], fn: ConsolidateFn): (number | null)[][] {
  const h = Math.max(0, ...sources.map((s) => s.data.length))
  const w = Math.max(0, ...sources.map((s) => Math.max(0, ...s.data.map((r) => r.length))))
  return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => combine(fn, sources.map((s) => s.data[i]?.[j] ?? null))))
}

/**
 * By labels: rows are keyed by the first column (when `left`) and columns by the first row (when `top`); matching
 * labels are combined across sources. Returns the table including its label row/column.
 */
export function consolidateByLabels(sources: Source[], fn: ConsolidateFn, top: boolean, left: boolean): Cell[][] {
  const rowKeys: string[] = [], colKeys: string[] = []
  const seenRow = new Set<string>(), seenCol = new Set<string>()
  const bucket = new Map<string, Cell[]>()
  for (const s of sources) {
    const heads = top ? s.data[0] ?? [] : []
    s.data.forEach((row, i) => {
      if (top && i === 0) return
      const rk = left ? String(row[0] ?? '') : String(i)
      if (!seenRow.has(rk.toLowerCase())) { seenRow.add(rk.toLowerCase()); rowKeys.push(rk) }
      row.forEach((v, j) => {
        if (left && j === 0) return
        const ck = top ? String(heads[j] ?? '') : String(j)
        if (!seenCol.has(ck.toLowerCase())) { seenCol.add(ck.toLowerCase()); colKeys.push(ck) }
        const k = rk.toLowerCase() + '\u0000' + ck.toLowerCase()
        let b = bucket.get(k)
        if (!b) bucket.set(k, (b = []))
        b.push(v)
      })
    })
  }
  const out: Cell[][] = []
  if (top) out.push([...(left ? [''] : []), ...colKeys])
  for (const rk of rowKeys) {
    out.push([...(left ? [rk] : []), ...colKeys.map((ck) => combine(fn, bucket.get(rk.toLowerCase() + '\u0000' + ck.toLowerCase()) ?? []))])
  }
  return out
}

// Pure maths behind the Statistic / Hierarchy / Surface charts (kept out of the SVG code so it can be tested).

const nums = (vals: (number | null)[]): number[] => vals.filter((v): v is number => typeof v === 'number' && isFinite(v))

/** Excel's QUARTILE.EXC-style quantile (p in 0..1); falls back to the nearest end when p is outside the exclusive range. */
export function quantileExc(sorted: number[], p: number): number {
  const n = sorted.length
  if (n === 0) return NaN
  const pos = p * (n + 1)
  if (pos <= 1) return sorted[0]
  if (pos >= n) return sorted[n - 1]
  const lo = Math.floor(pos)
  return sorted[lo - 1] + (pos - lo) * (sorted[lo] - sorted[lo - 1])
}

export interface BoxStats { q1: number; median: number; q3: number; mean: number; whiskerLo: number; whiskerHi: number; outliers: number[]; n: number }

/** Box & Whisker: exclusive-median quartiles, whiskers reach the furthest points within 1.5 x IQR, others are outliers. */
export function boxStats(vals: (number | null)[]): BoxStats | null {
  const v = nums(vals).sort((a, b) => a - b)
  if (!v.length) return null
  const q1 = quantileExc(v, 0.25), median = quantileExc(v, 0.5), q3 = quantileExc(v, 0.75)
  const iqr = q3 - q1
  const inside = v.filter((x) => x >= q1 - 1.5 * iqr && x <= q3 + 1.5 * iqr)
  return { q1, median, q3, mean: v.reduce((a, b) => a + b, 0) / v.length, whiskerLo: inside[0], whiskerHi: inside[inside.length - 1], outliers: v.filter((x) => x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr), n: v.length }
}

export interface HistBin { lo: number; hi: number; count: number; label: string }

const niceNum = (x: number): string => String(Math.round(x * 1e6) / 1e6)

/** Histogram bins using Excel's automatic bin width (Scott's rule: 3.5 sigma / n^(1/3)); first bin is [lo, hi], the rest (lo, hi]. */
export function histogramBins(vals: (number | null)[], width?: number): HistBin[] {
  const v = nums(vals)
  if (!v.length) return []
  const min = Math.min(...v), max = Math.max(...v)
  if (min === max) return [{ lo: min, hi: max, count: v.length, label: `[${niceNum(min)}, ${niceNum(max)}]` }]
  const mean = v.reduce((a, b) => a + b, 0) / v.length
  const sd = Math.sqrt(v.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, v.length - 1))
  const w = width && width > 0 ? width : Math.max((3.5 * sd) / Math.cbrt(v.length), (max - min) / 200)
  const count = Math.min(200, Math.max(1, Math.ceil((max - min) / w - 1e-9)))
  const bins: HistBin[] = Array.from({ length: count }, (_, i) => {
    const lo = min + i * w, hi = i === count - 1 ? Math.max(max, min + (i + 1) * w) : min + (i + 1) * w
    return { lo, hi, count: 0, label: `${i === 0 ? '[' : '('}${niceNum(lo)}, ${niceNum(hi)}]` }
  })
  for (const x of v) {
    let i = Math.min(count - 1, Math.max(0, Math.ceil((x - min) / w - 1e-9) - 1))
    if (x === min) i = 0
    bins[i].count++
  }
  return bins
}

/** Pareto: bins sorted by descending count with a running cumulative percentage. */
export function paretoOf(bins: HistBin[]): { label: string; count: number; cumPct: number }[] {
  const sorted = [...bins].sort((a, b) => b.count - a.count)
  const total = sorted.reduce((a, b) => a + b.count, 0) || 1
  let run = 0
  return sorted.map((b) => ({ label: b.label, count: b.count, cumPct: (run += b.count) / total * 100 }))
}

export interface Tile { i: number; x: number; y: number; w: number; h: number }

/** Squarified treemap of `values` (non-positive entries get no tile) inside the box. */
export function squarify(values: number[], box: { x: number; y: number; w: number; h: number }): Tile[] {
  const items = values.map((v, i) => ({ i, v })).filter((t) => t.v > 0).sort((a, b) => b.v - a.v)
  const total = items.reduce((a, b) => a + b.v, 0)
  const out: Tile[] = []
  if (!total) return out
  const scale = (box.w * box.h) / total
  let { x, y, w, h } = box
  let row: { i: number; a: number }[] = []
  const worst = (r: typeof row, side: number): number => {
    const s = r.reduce((a, b) => a + b.a, 0)
    const mx = Math.max(...r.map((t) => t.a)), mn = Math.min(...r.map((t) => t.a))
    return Math.max((side * side * mx) / (s * s), (s * s) / (side * side * mn))
  }
  const flush = (): void => {
    if (!row.length) return
    const s = row.reduce((a, b) => a + b.a, 0)
    if (w >= h) { // lay the row as a column on the left
      const cw = s / h; let cy = y
      for (const t of row) { out.push({ i: t.i, x, y: cy, w: cw, h: t.a / cw }); cy += t.a / cw }
      x += cw; w -= cw
    } else {
      const rh = s / w; let cx = x
      for (const t of row) { out.push({ i: t.i, x: cx, y, w: t.a / rh, h: rh }); cx += t.a / rh }
      y += rh; h -= rh
    }
    row = []
  }
  for (const it of items) {
    const a = it.v * scale
    const side = Math.min(w, h)
    if (row.length && worst([...row, { i: it.i, a }], side) > worst(row, side)) flush()
    row.push({ i: it.i, a })
  }
  flush()
  return out
}

/** Surface bands: n equal-width value ranges over [min, max], used for colouring and the legend. */
export function surfaceBands(min: number, max: number, count = 6): { lo: number; hi: number }[] {
  const span = (max - min) || 1
  return Array.from({ length: count }, (_, i) => ({ lo: min + (span * i) / count, hi: min + (span * (i + 1)) / count }))
}

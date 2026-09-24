// Row/column offset lookups that stay fast with up to MAX_ROWS/MAX_COLS indices while only a
// sparse set of rows/columns have a custom size or are hidden.
export interface Axis {
  /** Pixel offset of the start of index i (from index 0). */
  offset(i: number): number
  /** Size in px of index i (0 if hidden). */
  size(i: number): number
  /** Largest index whose offset is <= px. */
  indexAt(px: number): number
  /** Total pixel size of the first `n` indices. */
  totalSize(n: number): number
}

export function buildAxis(overrides: Record<number, number>, hidden: Record<number, true>, def: number, count: number): Axis {
  const keys = new Set<number>()
  for (const k in overrides) keys.add(+k)
  for (const k in hidden) keys.add(+k)
  const sorted = [...keys].sort((a, b) => a - b).filter((i) => i < count)
  const cumDelta: number[] = new Array(sorted.length)
  let cum = 0
  for (let i = 0; i < sorted.length; i++) {
    const idx = sorted[i]
    const h = hidden[idx] ? 0 : (overrides[idx] ?? def)
    cum += h - def
    cumDelta[i] = cum
  }

  function deltaBefore(i: number): number {
    // sum of deltas for overridden indices < i
    let lo = 0, hi = sorted.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (sorted[mid] < i) lo = mid + 1
      else hi = mid
    }
    return lo > 0 ? cumDelta[lo - 1] : 0
  }

  function offset(i: number): number {
    if (i <= 0) return 0
    return i * def + deltaBefore(Math.min(i, count))
  }
  function size(i: number): number {
    if (hidden[i]) return 0
    return overrides[i] ?? def
  }
  function indexAt(px: number): number {
    if (px <= 0) return 0
    let lo = 0, hi = Math.max(0, count - 1)
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (offset(mid) <= px) lo = mid
      else hi = mid - 1
    }
    return lo
  }
  function totalSize(n: number): number {
    return offset(Math.min(n, count))
  }
  return { offset, size, indexAt, totalSize }
}

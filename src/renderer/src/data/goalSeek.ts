// What-If Analysis > Goal Seek: finds the input value that makes a formula cell reach a target.
// `f` recomputes the formula cell for a candidate input (the caller writes the input and reads the result).
export interface GoalSeekResult { found: boolean; x: number; value: number; iterations: number }

export function goalSeek(f: (x: number) => number | null, target: number, start: number, maxIter = 100, tol = 0.001): GoalSeekResult {
  const g = (x: number): number | null => { const v = f(x); return v === null || !isFinite(v) ? null : v - target }
  let x = start
  let gx = g(x)
  if (gx === null) return { found: false, x, value: NaN, iterations: 0 }
  let it = 0
  // Secant iteration with an adaptive first step, then bracketing/bisection if it stalls.
  let x0 = x, g0 = gx
  let x1 = x + (Math.abs(x) > 1e-9 ? Math.abs(x) * 0.01 : 0.01)
  let g1 = g(x1)
  while (it < maxIter) {
    if (Math.abs(gx) <= tol) return { found: true, x, value: gx + target, iterations: it }
    if (g1 === null) break
    if (Math.abs(g1) <= tol) return { found: true, x: x1, value: g1 + target, iterations: it }
    if (g1 === g0) { x1 += (x1 - x0) || 1; g1 = g(x1); it++; continue }
    const x2 = x1 - (g1 * (x1 - x0)) / (g1 - g0)
    x0 = x1; g0 = g1
    x1 = x2; g1 = g(x1)
    x = x1; gx = g1 ?? gx
    it++
  }
  // Fallback: expand a bracket around the start and bisect.
  let lo = start, hi = start, glo = g(lo), ghi = glo
  let step = Math.max(1, Math.abs(start))
  for (let k = 0; k < 60 && glo !== null && ghi !== null && Math.sign(glo) === Math.sign(ghi); k++) {
    lo = start - step; hi = start + step
    glo = g(lo); ghi = g(hi)
    step *= 2
  }
  if (glo === null || ghi === null || Math.sign(glo) === Math.sign(ghi)) return { found: false, x: start, value: (gx ?? NaN) + target, iterations: it }
  for (let k = 0; k < 200; k++) {
    const mid = (lo + hi) / 2
    const gm = g(mid)
    if (gm === null) break
    if (Math.abs(gm) <= tol) return { found: true, x: mid, value: gm + target, iterations: it + k }
    if (Math.sign(gm) === Math.sign(glo)) { lo = mid; glo = gm } else { hi = mid; ghi = gm }
  }
  return { found: false, x: (lo + hi) / 2, value: NaN, iterations: it }
}

// Pure logic behind the Combo and Waterfall charts (no React), shared by the renderer and the .xlsx writer.
import { ChartSpec } from '../model/types'

export type ComboSeries = NonNullable<ChartSpec['comboSeries']>[number]

/** Excel's default combo: every series a clustered column except the last, which is a line (optionally on a secondary axis). */
export function defaultComboSeries(count: number, secondary: boolean): ComboSeries[] {
  return Array.from({ length: Math.max(1, count) }, (_, i) => (i === count - 1 && count > 1 ? { type: 'line', ...(secondary ? { secondary: true } : {}) } : { type: 'bar' }))
}

/** The plot type of series `i`, falling back to the default combo when the spec has no (or a short) per-series list. */
export function comboSeriesOf(spec: Pick<ChartSpec, 'comboSeries'>, count: number): ComboSeries[] {
  const def = defaultComboSeries(count, false)
  return def.map((d, i) => spec.comboSeries?.[i] ?? d)
}

export interface WaterfallStep { start: number; end: number; kind: 'up' | 'down' }

/** Running-total steps for a waterfall: each value moves the total from `start` to `end`. Blank values are zero. */
export function waterfallSteps(values: (number | null)[]): WaterfallStep[] {
  let total = 0
  return values.map((v) => {
    const d = v ?? 0
    const start = total
    total += d
    return { start, end: total, kind: d >= 0 ? 'up' : 'down' }
  })
}

/** Increase / Decrease / invisible-base columns that draw a waterfall as a stacked column chart (the .xlsx form). */
export function waterfallStack(values: (number | null)[]): { base: number[]; increase: number[]; decrease: number[] } {
  const base: number[] = [], increase: number[] = [], decrease: number[] = []
  for (const s of waterfallSteps(values)) {
    base.push(Math.min(s.start, s.end))
    increase.push(s.kind === 'up' ? s.end - s.start : 0)
    decrease.push(s.kind === 'down' ? s.start - s.end : 0)
  }
  return { base, increase, decrease }
}

/** A "nice" 1/2/5 x 10^k axis spanning [min, max] that always includes zero: gridline values low..high. */
export function niceRange(min: number, max: number): { lo: number; hi: number; step: number } {
  const lo0 = Math.min(0, min), hi0 = Math.max(0, max)
  const span = hi0 - lo0 || 1
  const raw = span / 5
  const mag = Math.pow(10, Math.floor(Math.log10(raw)))
  const norm = raw / mag
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag
  return { lo: Math.floor(lo0 / step + 1e-9) * step, hi: Math.ceil(hi0 / step - 1e-9) * step, step }
}

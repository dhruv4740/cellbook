import { describe, it, expect } from 'vitest'
import { axisScale, axisScaleRange } from './ChartRenderer'

describe('axisScaleRange', () => {
  it('matches axisScale for all-positive data', () => {
    for (const v of [0, 0.3, 1, 5, 7, 42, 100, 980, 12345]) {
      const a = axisScale(v), r = axisScaleRange(0, v)
      expect(r.min).toBe(0)
      expect(r.max).toBe(a.max)
      expect(r.ticks).toBe(a.ticks)
      expect(r.step * r.ticks).toBeCloseTo(a.max)
    }
  })
  it('spans negative and positive data (Excel -40..120 step 20)', () => {
    expect(axisScaleRange(-30, 100)).toEqual({ min: -40, max: 120, ticks: 8, step: 20 })
  })
  it('ends at zero for negative-only data', () => {
    expect(axisScaleRange(-30, 0)).toEqual({ min: -35, max: 0, ticks: 7, step: 5 })
  })
  it('keeps ticks integral and the axis covering the data', () => {
    for (const [lo, hi] of [[-1, 1], [-7, 3], [-250, 900], [-0.4, 0.9]]) {
      const r = axisScaleRange(lo, hi)
      expect(r.min).toBeLessThan(lo)
      expect(r.max).toBeGreaterThan(hi)
      expect(Number.isInteger(r.ticks)).toBe(true)
      expect((r.max - r.min) / r.step).toBeCloseTo(r.ticks)
    }
  })
})

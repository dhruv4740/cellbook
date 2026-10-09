import { describe, it, expect } from 'vitest'
import { boxStats, histogramBins, paretoOf, quantileExc, squarify, surfaceBands } from './extraStats'

describe('extraStats', () => {
  it('quartiles match QUARTILE.EXC', () => {
    const v = [1, 2, 3, 4, 5, 6, 7, 8]
    expect(quantileExc(v, 0.25)).toBe(2.25)
    expect(quantileExc(v, 0.5)).toBe(4.5)
    expect(quantileExc(v, 0.75)).toBe(6.75)
  })
  it('box stats flag outliers beyond 1.5 IQR', () => {
    const b = boxStats([1, 2, 3, 4, 5, 6, 7, 8, 100])!
    expect(b.outliers).toEqual([100])
    expect(b.whiskerHi).toBe(8)
    expect(b.median).toBe(5)
  })
  it('histogram counts every value once and uses [a,b] then (a,b] labels', () => {
    const vals = [1, 2, 2, 3, 3, 3, 4, 4, 5, 9]
    const bins = histogramBins(vals, 2)
    expect(bins.reduce((a, b) => a + b.count, 0)).toBe(vals.length)
    expect(bins[0].label).toBe('[1, 3]')
    expect(bins[1].label).toBe('(3, 5]')
    expect(bins[0].count).toBe(6) // 1,2,2,3,3,3
  })
  it('pareto sorts descending and ends at 100%', () => {
    const p = paretoOf(histogramBins([1, 2, 2, 3, 3, 3, 4, 4, 5, 9], 2))
    expect(p[0].count).toBeGreaterThanOrEqual(p[1].count)
    expect(p[p.length - 1].cumPct).toBeCloseTo(100)
  })
  it('treemap tiles tile the box area in proportion to the values', () => {
    const tiles = squarify([6, 6, 4, 3, 2, 2, 1], { x: 0, y: 0, w: 100, h: 60 })
    expect(tiles.length).toBe(7)
    expect(tiles.reduce((a, t) => a + t.w * t.h, 0)).toBeCloseTo(6000)
    const t0 = tiles.find((t) => t.i === 0)!
    expect(t0.w * t0.h).toBeCloseTo(6000 * 6 / 24)
  })
  it('surface bands split the range evenly', () => {
    const b = surfaceBands(0, 60, 6)
    expect(b[0]).toEqual({ lo: 0, hi: 10 })
    expect(b[5].hi).toBe(60)
  })
})

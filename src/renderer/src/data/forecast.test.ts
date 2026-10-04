import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { buildForecastGrid, createForecastSheet, parseForecastEnd, readForecastSource } from './forecast'

describe('Data > Forecast Sheet', () => {
  it('parses the Forecast End box', () => {
    expect(parseForecastEnd('12')).toBe(12)
    expect(parseForecastEnd('2026-01-01')).toBe(46023)
    expect(parseForecastEnd('nope')).toBeNull()
  })
  it('builds the grid with live formulas', () => {
    const g = buildForecastGrid({ xs: [1, 2, 3], ys: [2, 4, 6], xHead: 'T', yHead: 'V' }, { chart: 'line', end: 5, confidence: 0.95 })
    expect(g.length).toBe(6)
    expect(g[0].slice(0, 5)).toEqual(['T', 'V', 'Forecast(V)', 'Lower Confidence Bound(V)', 'Upper Confidence Bound(V)'])
    expect(g[4][2]).toBe('=FORECAST.LINEAR(A5,$B$2:$B$4,$A$2:$A$4)')
    expect(g[1][7]).toBe('1')
  })
  it('creates a sheet whose forecast and bounds evaluate, plus a chart; undo removes it', () => {
    const s = new WorkbookStore()
    const sh = s.active
    ;[['T', 'V'], ['1', '2.1'], ['2', '3.9'], ['3', '6.2'], ['4', '7.8']].forEach((r, i) => r.forEach((v, j) => s.setRaw(sh, i, j, v)))
    const src = readForecastSource(s, sh, { r1: 0, c1: 0, r2: 4, c2: 1 })
    if (typeof src === 'string') throw new Error(src)
    expect(createForecastSheet(s, src, { chart: 'line', end: 6, confidence: 0.95 })).toBeNull()
    const f = s.active
    expect(f).not.toBe(sh)
    expect(Number(s.getValue(f, 5, 2))).toBeCloseTo(10.0, 0)
    expect(Number(s.getValue(f, 5, 3))).toBeLessThan(Number(s.getValue(f, 5, 2)))
    expect(Number(s.getValue(f, 5, 4))).toBeGreaterThan(Number(s.getValue(f, 5, 2)))
    expect(f.layout.charts?.length).toBe(1)
    s.undo()
    expect(s.sheets.length).toBe(1)
  })
})

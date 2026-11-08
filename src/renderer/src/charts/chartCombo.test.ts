import { describe, it, expect } from 'vitest'
import { defaultComboSeries, comboSeriesOf, waterfallSteps, waterfallStack, niceRange } from './chartCombo'

describe('combo series', () => {
  it('defaults to columns with the last series as a line', () => {
    expect(defaultComboSeries(3, false)).toEqual([{ type: 'bar' }, { type: 'bar' }, { type: 'line' }])
    expect(defaultComboSeries(2, true)).toEqual([{ type: 'bar' }, { type: 'line', secondary: true }])
    expect(defaultComboSeries(1, false)).toEqual([{ type: 'bar' }])
  })
  it('uses the spec list and falls back for missing entries', () => {
    expect(comboSeriesOf({ comboSeries: [{ type: 'line' }] }, 2)).toEqual([{ type: 'line' }, { type: 'line' }])
    expect(comboSeriesOf({}, 2)).toEqual([{ type: 'bar' }, { type: 'line' }])
  })
})

describe('waterfall', () => {
  it('accumulates a running total, treating blanks as zero', () => {
    expect(waterfallSteps([100, -30, null, 50])).toEqual([
      { start: 0, end: 100, kind: 'up' }, { start: 100, end: 70, kind: 'down' },
      { start: 70, end: 70, kind: 'up' }, { start: 70, end: 120, kind: 'up' }
    ])
  })
  it('builds the invisible-base stacked form, including bars that cross zero', () => {
    expect(waterfallStack([100, -30, -90])).toEqual({ base: [0, 70, -20], increase: [100, 0, 0], decrease: [0, 30, 90] })
  })
})

describe('niceRange', () => {
  it('spans zero and the data', () => {
    expect(niceRange(0, 120)).toEqual({ lo: 0, hi: 150, step: 50 })
    const r = niceRange(-20, 100)
    expect(r.lo).toBeLessThanOrEqual(-20)
    expect(r.hi).toBeGreaterThanOrEqual(100)
  })
})

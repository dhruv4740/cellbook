import { describe, it, expect } from 'vitest'
import { seriesValues } from './series'

describe('Fill Series', () => {
  it('linear with step and stop', () => {
    expect(seriesValues(1, 10, 'linear', 2, 7)).toEqual([3, 5, 7])
    expect(seriesValues(0.1, 3, 'linear', 0.1)).toEqual([0.2, 0.3, 0.4])
  })
  it('growth', () => {
    expect(seriesValues(2, 3, 'growth', 3)).toEqual([6, 18, 54])
  })
  it('dates by day, weekday, month, year', () => {
    // 45688 = 31-Jan-2025 (Friday)
    expect(seriesValues(45688, 2, 'date', 1, undefined, 'day')).toEqual([45689, 45690])
    expect(seriesValues(45688, 2, 'date', 1, undefined, 'weekday')).toEqual([45691, 45692]) // Mon 3 Feb, Tue 4 Feb
    expect(seriesValues(45688, 2, 'date', 1, undefined, 'month')).toEqual([45716, 45747]) // 28 Feb, 31 Mar
    expect(seriesValues(45688, 1, 'date', 1, undefined, 'year')).toEqual([46053])
  })
})

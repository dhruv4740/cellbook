import { describe, it, expect } from 'vitest'
import { generateSeries } from './autofill'

describe('generateSeries', () => {
  it('copies a single plain value', () => {
    expect(generateSeries(['5'], 3)).toEqual(['5', '5', '5'])
  })
  it('extends a linear number trend from 2 seeds', () => {
    expect(generateSeries(['1', '2'], 4)).toEqual(['3', '4', '5', '6'])
  })
  it('extends a linear trend from more seeds (best fit)', () => {
    expect(generateSeries(['2', '4', '6'], 2)).toEqual(['8', '10'])
  })
  it('increments a trailing number in text', () => {
    expect(generateSeries(['Item 1'], 2)).toEqual(['Item 2', 'Item 3'])
  })
  it('keeps zero-padding width on trailing numbers', () => {
    expect(generateSeries(['Item 001'], 2)).toEqual(['Item 002', 'Item 003'])
  })
  it('cycles short weekday names', () => {
    expect(generateSeries(['Mon'], 3)).toEqual(['Tue', 'Wed', 'Thu'])
  })
  it('cycles weekday names across week boundary', () => {
    expect(generateSeries(['Fri'], 3)).toEqual(['Sat', 'Sun', 'Mon'])
  })
  it('cycles long month names', () => {
    expect(generateSeries(['November'], 3)).toEqual(['December', 'January', 'February'])
  })
  it('increments dates by day for a single date seed', () => {
    expect(generateSeries(['1/1/2024'], 2)).toEqual(['1/2/2024', '1/3/2024'])
  })
  it('extends date step from two seeds', () => {
    expect(generateSeries(['1/1/2024', '1/3/2024'], 2)).toEqual(['1/5/2024', '1/7/2024'])
  })
  it('copy mode repeats the seed list verbatim, ignoring patterns', () => {
    expect(generateSeries(['1', '2'], 4, true)).toEqual(['1', '2', '1', '2'])
  })
})

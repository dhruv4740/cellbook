import { describe, it, expect } from 'vitest'
import { distinctColumnValues, computeFilteredRows, toggleValue, BLANK_TOKEN } from './filter'
import { Range, FilterCriteria } from '../model/types'

// Grid: header on row 0, data rows 1..4, single column 0.
const data = ['Fruit', 'Apple', 'Banana', '', 'Apple']
const range: Range = { r1: 0, c1: 0, r2: 4, c2: 0 }
const getDisplay = (r: number, c: number): string => (c === 0 ? data[r] : '')

describe('distinctColumnValues', () => {
  it('collects unique values below the header, blanks as a sentinel', () => {
    const vals = distinctColumnValues(range, 0, getDisplay)
    expect(vals.sort()).toEqual(['Apple', 'Banana', BLANK_TOKEN].sort())
  })
})

describe('computeFilteredRows', () => {
  it('shows all rows when every value is selected', () => {
    const columns: Record<number, FilterCriteria> = { 0: { selected: ['Apple', 'Banana', BLANK_TOKEN] } }
    const { hidden, visible, total } = computeFilteredRows(range, columns, getDisplay)
    expect(hidden).toEqual([])
    expect(visible).toBe(4)
    expect(total).toBe(4)
  })

  it('hides rows whose value is unchecked', () => {
    const columns: Record<number, FilterCriteria> = { 0: { selected: ['Apple'] } }
    const { hidden, visible } = computeFilteredRows(range, columns, getDisplay)
    expect(hidden).toEqual([2, 3]) // Banana row, blank row
    expect(visible).toBe(2)
  })

  it('can filter to just blanks', () => {
    const columns: Record<number, FilterCriteria> = { 0: { selected: [BLANK_TOKEN] } }
    const { hidden, visible } = computeFilteredRows(range, columns, getDisplay)
    expect(visible).toBe(1)
    expect(hidden).toEqual([1, 2, 4])
  })

  it('combines multiple column criteria with AND', () => {
    const data2 = [
      ['A', 'B'],
      ['x', '1'],
      ['x', '2'],
      ['y', '1']
    ]
    const rg: Range = { r1: 0, c1: 0, r2: 3, c2: 1 }
    const get2 = (r: number, c: number) => data2[r][c]
    const columns: Record<number, FilterCriteria> = { 0: { selected: ['x'] }, 1: { selected: ['1'] } }
    const { hidden, visible } = computeFilteredRows(rg, columns, get2)
    expect(visible).toBe(1)
    expect(hidden).toEqual([2, 3])
  })

  it('an empty range has nothing to hide', () => {
    const rg: Range = { r1: 0, c1: 0, r2: 0, c2: 0 }
    const { hidden, visible, total } = computeFilteredRows(rg, {}, getDisplay)
    expect(hidden).toEqual([])
    expect(visible).toBe(0)
    expect(total).toBe(0)
  })
})

describe('toggleValue', () => {
  const all = ['Apple', 'Banana', BLANK_TOKEN]
  it('removes a checked value', () => {
    expect(toggleValue(all, 'Apple', all)).toEqual(['Banana', BLANK_TOKEN])
  })
  it('adds back an unchecked value, preserving order', () => {
    expect(toggleValue(['Banana'], 'Apple', all)).toEqual(['Apple', 'Banana'])
  })
})

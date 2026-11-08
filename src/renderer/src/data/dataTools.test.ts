import { describe, expect, it } from 'vitest'
import { advancedFilterRows, matchesCriterion, uniqueRows } from './advancedFilter'
import { goalSeek } from './goalSeek'
import { consolidateByLabels, consolidateByPosition } from './consolidate'

const cell = (v: string | number | null) => ({ text: v === null ? '' : String(v), value: v })

describe('advanced filter criteria', () => {
  it('numbers and operators', () => {
    expect(matchesCriterion('>5', cell(6))).toBe(true)
    expect(matchesCriterion('>5', cell(5))).toBe(false)
    expect(matchesCriterion('<=5', cell(5))).toBe(true)
    expect(matchesCriterion('5', cell(5))).toBe(true)
    expect(matchesCriterion('<>5', cell(4))).toBe(true)
    expect(matchesCriterion('>5', cell('abc'))).toBe(false)
  })
  it('text begins-with, exact and wildcards', () => {
    expect(matchesCriterion('Da', cell('Dave'))).toBe(true)
    expect(matchesCriterion('=Da', cell('Dave'))).toBe(false)
    expect(matchesCriterion('=dave', cell('Dave'))).toBe(true)
    expect(matchesCriterion('*ve', cell('Dave'))).toBe(true)
    expect(matchesCriterion('D?ve', cell('Dave'))).toBe(true)
    expect(matchesCriterion('<>Dave', cell('Dave'))).toBe(false)
    expect(matchesCriterion('=', cell(null))).toBe(true)
    expect(matchesCriterion('<>', cell(null))).toBe(false)
  })
  it('ANDs across a row and ORs across rows', () => {
    const data = [['East', 10], ['West', 20], ['East', 30], ['North', 5]]
    const read = (r: number, c: number) => cell(data[r][c] as string | number)
    const heads = ['Region', 'Sales']
    expect(advancedFilterRows(heads, 4, read, [['Region', 'Sales'], ['East', '>15']])).toEqual([2])
    expect(advancedFilterRows(heads, 4, read, [['Region', 'Sales'], ['East', ''], ['West', '']])).toEqual([0, 1, 2])
    expect(advancedFilterRows(heads, 4, read, [['Region']])).toEqual([0, 1, 2, 3])
  })
  it('unique rows', () => {
    expect(uniqueRows([0, 1, 2], (r) => ['a', 'A', 'b'][r])).toEqual([0, 2])
  })
})

describe('goal seek', () => {
  it('solves a linear equation', () => {
    const r = goalSeek((x) => 3 * x + 2, 11, 0)
    expect(r.found).toBe(true)
    expect(r.x).toBeCloseTo(3, 2)
  })
  it('solves a nonlinear one', () => {
    const r = goalSeek((x) => x * x, 49, 1)
    expect(r.found).toBe(true)
    expect(Math.abs(r.x)).toBeCloseTo(7, 2)
  })
  it('reports failure when unreachable', () => {
    expect(goalSeek((x) => x * x + 1, -5, 1).found).toBe(false)
  })
})

describe('consolidate', () => {
  it('by position', () => {
    const r = consolidateByPosition([{ data: [[1, 2], [3, 4]] }, { data: [[10, 20], [30, 40]] }], 'Sum')
    expect(r).toEqual([[11, 22], [33, 44]])
  })
  it('by labels', () => {
    const a = { data: [['', 'Q1', 'Q2'], ['x', 1, 2], ['y', 3, 4]] }
    const b = { data: [['', 'Q1', 'Q2'], ['y', 10, 20], ['z', 5, 6]] }
    expect(consolidateByLabels([a, b], 'Sum', true, true)).toEqual([['', 'Q1', 'Q2'], ['x', 1, 2], ['y', 13, 24], ['z', 5, 6]])
    expect(consolidateByLabels([a, b], 'Max', true, true)[2]).toEqual(['y', 10, 20])
  })
})

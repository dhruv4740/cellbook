import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { compareValues, sortRange, detectHeaders, findCurrentRegion, hasAdjacentData } from './sort'

describe('compareValues matching Excel hierarchy', () => {
  it('ranks numbers < text < logicals < errors < blanks in ascending order', () => {
    const list = ['text', 10, true, '', { error: '#VALUE!' }, 5, false, 'abc', null]
    const sorted = [...list].sort((a, b) => compareValues(a, b, 'asc'))
    // Numbers: 5, 10
    // Text: 'abc', 'text'
    // Logicals: false, true
    // Errors: #VALUE!
    // Blanks: '', null
    expect(sorted[0]).toBe(5)
    expect(sorted[1]).toBe(10)
    expect(sorted[2]).toBe('abc')
    expect(sorted[3]).toBe('text')
    expect(sorted[4]).toBe(false)
    expect(sorted[5]).toBe(true)
    expect((sorted[6] as any).error).toBe('#VALUE!')
    // Blanks at the very end
    expect(sorted[7] === '' || sorted[7] === null).toBe(true)
    expect(sorted[8] === '' || sorted[8] === null).toBe(true)
  })

  it('keeps blanks last in descending order', () => {
    const list = [10, '', 5, null, 20]
    const sorted = [...list].sort((a, b) => compareValues(a, b, 'desc'))
    expect(sorted[0]).toBe(20)
    expect(sorted[1]).toBe(10)
    expect(sorted[2]).toBe(5)
    expect(sorted[3] === '' || sorted[3] === null).toBe(true)
    expect(sorted[4] === '' || sorted[4] === null).toBe(true)
  })
})

describe('detectHeaders', () => {
  it('detects string header row over numeric data', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 0, raw: 'Name' },
      { r: 0, c: 1, raw: 'Score' },
      { r: 1, c: 0, raw: 'Alice' },
      { r: 1, c: 1, raw: '95' }
    ], 'Test')
    expect(detectHeaders(store, sheet, { r1: 0, c1: 0, r2: 1, c2: 1 })).toBe(true)
  })

  it('returns false for pure numbers', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 0, raw: '10' },
      { r: 1, c: 0, raw: '20' }
    ], 'Test')
    expect(detectHeaders(store, sheet, { r1: 0, c1: 0, r2: 1, c2: 0 })).toBe(false)
  })
})

describe('sortRange', () => {
  it('sorts multi-level criteria with formula shift', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 0, raw: 'Category' },
      { r: 0, c: 1, raw: 'Val' },
      { r: 0, c: 2, raw: 'Formula' },
      { r: 1, c: 0, raw: 'B' },
      { r: 1, c: 1, raw: '20' },
      { r: 1, c: 2, raw: '=B2*2' },
      { r: 2, c: 0, raw: 'A' },
      { r: 2, c: 1, raw: '50' },
      { r: 2, c: 2, raw: '=B3*2' },
      { r: 3, c: 0, raw: 'A' },
      { r: 3, c: 1, raw: '10' },
      { r: 3, c: 2, raw: '=B4*2' }
    ], 'Test')

    sortRange(store, sheet, { r1: 0, c1: 0, r2: 3, c2: 2 }, {
      hasHeaders: true,
      levels: [
        { columnIndex: 0, sortOn: 'values', order: 'asc' },
        { columnIndex: 1, sortOn: 'values', order: 'asc' }
      ]
    })

    // After sort:
    // Header stays: Category, Val, Formula
    // Row 1: A, 10, formula shifted to =B2*2
    // Row 2: A, 50, formula shifted to =B3*2
    // Row 3: B, 20, formula shifted to =B4*2
    expect(store.getRaw(sheet, 1, 0)).toBe('A')
    expect(store.getValue(sheet, 1, 1)).toBe(10)
    expect(store.getRaw(sheet, 1, 2)).toBe('=B2*2')

    expect(store.getRaw(sheet, 2, 0)).toBe('A')
    expect(store.getValue(sheet, 2, 1)).toBe(50)
    expect(store.getRaw(sheet, 2, 2)).toBe('=B3*2')

    expect(store.getRaw(sheet, 3, 0)).toBe('B')
    expect(store.getValue(sheet, 3, 1)).toBe(20)
    expect(store.getRaw(sheet, 3, 2)).toBe('=B4*2')
  })
})

describe('findCurrentRegion and hasAdjacentData', () => {
  it('finds contiguous region and detects adjacent data', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 2, c: 2, raw: 'A' },
      { r: 2, c: 3, raw: 'B' },
      { r: 3, c: 2, raw: 'C' },
      { r: 3, c: 3, raw: 'D' }
    ], 'Test')

    const region = findCurrentRegion(store, sheet, { r1: 2, c1: 2, r2: 2, c2: 2 })
    expect(region).toEqual({ r1: 2, c1: 2, r2: 3, c2: 3 })

    expect(hasAdjacentData(store, sheet, { r1: 2, c1: 2, r2: 2, c2: 2 })).toBe(true)
    expect(hasAdjacentData(store, sheet, { r1: 2, c1: 2, r2: 3, c2: 3 })).toBe(false)
  })
})

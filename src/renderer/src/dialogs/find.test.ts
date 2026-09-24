import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import {
  compileWildcardPattern,
  cellMatches,
  findAll,
  findNext,
  replaceAll,
  getSpecialRanges
} from './find'

describe('wildcard matching', () => {
  it('handles *, ?, and ~ escapes properly', () => {
    expect(cellMatches('Hello World', 'Hello*', false, false)).toBe(true)
    expect(cellMatches('Hello World', 'world', false, false)).toBe(true)
    expect(cellMatches('Hello World', 'world', true, false)).toBe(false)
    expect(cellMatches('Hello World', 'Hello', false, true)).toBe(false)
    expect(cellMatches('Hello', 'H?llo', false, true)).toBe(true)
    expect(cellMatches('Price * 10', '*~**', false, false)).toBe(true)
    expect(cellMatches('What?', '*~?', false, true)).toBe(true)
  })
})

describe('search & replace in workbook', () => {
  it('finds in active sheet and workbook', () => {
    const store = new WorkbookStore()
    const s1 = store.active
    store.setRaw(s1, 0, 0, 'Apple')
    store.setRaw(s1, 0, 1, 'Banana')
    store.setRaw(s1, 1, 0, '=SUM(10, 20)')
    store.setRaw(s1, 1, 1, 'Pineapple')

    const res1 = findAll(store, { query: 'Apple', within: 'sheet' })
    expect(res1.length).toBe(2)
    expect(res1[0].cellRef).toBe('$A$1')
    expect(res1[1].cellRef).toBe('$B$2')

    // Formula search
    const resFormulas = findAll(store, { query: 'SUM', lookIn: 'formulas' })
    expect(resFormulas.length).toBe(1)
    expect(resFormulas[0].cellRef).toBe('$A$2')

    // Find next
    const next1 = findNext(store, { query: 'Apple' }, { sheetIndex: 0, r: 0, c: 0 })
    expect(next1?.cellRef).toBe('$B$2')
    // Wrapping
    const nextWrap = findNext(store, { query: 'Apple' }, { sheetIndex: 0, r: 1, c: 1 })
    expect(nextWrap?.cellRef).toBe('$A$1')
  })

  it('replaces all matches in single transaction', () => {
    const store = new WorkbookStore()
    const s1 = store.active
    store.setRaw(s1, 0, 0, 'cat')
    store.setRaw(s1, 0, 1, 'catch')
    store.setRaw(s1, 1, 0, 'scatter')

    const count = replaceAll(store, { query: 'cat', matchEntireCell: false }, 'dog')
    expect(count).toBe(3)
    expect(store.getRaw(s1, 0, 0)).toBe('dog')
    expect(store.getRaw(s1, 0, 1)).toBe('dogch')
    expect(store.getRaw(s1, 1, 0)).toBe('sdogter')

    // Undo restores all three
    store.undo()
    expect(store.getRaw(s1, 0, 0)).toBe('cat')
    expect(store.getRaw(s1, 0, 1)).toBe('catch')
    expect(store.getRaw(s1, 1, 0)).toBe('scatter')
  })
})

describe('Go To Special ranges', () => {
  it('selects constants, formulas, blanks, lastCell', () => {
    const store = new WorkbookStore()
    const s1 = store.active
    store.setRaw(s1, 0, 0, 'Text')
    store.setRaw(s1, 0, 1, '123')
    store.setRaw(s1, 1, 0, '=1+2')
    // row 1 col 1 is blank

    const constants = getSpecialRanges(store, 'constants')
    expect(constants.length).toBe(2) // A1 and B1

    const formulas = getSpecialRanges(store, 'formulas')
    expect(formulas.length).toBe(1) // A2

    const blanks = getSpecialRanges(store, 'blanks')
    expect(blanks.length).toBe(1) // B2

    const last = getSpecialRanges(store, 'lastCell')
    expect(last).toEqual([{ r1: 1, c1: 1, r2: 1, c2: 1 }])
  })
})

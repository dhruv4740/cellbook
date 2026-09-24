import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { splitDelimitedLine, parseFixedLine, convertTextToColumns } from './textToColumns'

describe('splitDelimitedLine', () => {
  it('splits by comma honoring quotes', () => {
    const line = 'Apple,"Banana, Yellow",Cherry'
    const parts = splitDelimitedLine(line, [','], false, '"')
    expect(parts).toEqual(['Apple', 'Banana, Yellow', 'Cherry'])
  })

  it('treats consecutive delimiters as one when configured', () => {
    const line = 'Word1   Word2     Word3'
    const parts = splitDelimitedLine(line, [' '], true, '"')
    expect(parts).toEqual(['Word1', 'Word2', 'Word3'])
  })
})

describe('parseFixedLine', () => {
  it('splits line at fixed break positions', () => {
    const line = '1234567890ABCDEF'
    const parts = parseFixedLine(line, [5, 10])
    expect(parts).toEqual(['12345', '67890', 'ABCDEF'])
  })
})

describe('convertTextToColumns', () => {
  it('splits delimited data into multiple columns starting at destination', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 0, raw: 'First,Last,Score' },
      { r: 1, c: 0, raw: 'John,Doe,88' },
      { r: 2, c: 0, raw: 'Jane,Smith,92' }
    ], 'Test')

    convertTextToColumns(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 0 }, {
      mode: 'delimited',
      delimiters: { comma: true },
      destination: '$B$1'
    })

    expect(store.getRaw(sheet, 0, 1)).toBe('First')
    expect(store.getRaw(sheet, 0, 2)).toBe('Last')
    expect(store.getRaw(sheet, 0, 3)).toBe('Score')

    expect(store.getRaw(sheet, 1, 1)).toBe('John')
    expect(store.getRaw(sheet, 1, 2)).toBe('Doe')
    expect(store.getValue(sheet, 1, 3)).toBe(88)
  })
})

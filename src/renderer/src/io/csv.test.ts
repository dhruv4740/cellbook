import { describe, it, expect } from 'vitest'
import { readCsv, writeCsv } from './csv'

describe('readCsv', () => {
  it('parses simple comma-separated rows', () => {
    const data = readCsv('a,b,c\r\n1,2,3\r\n')
    expect(data.sheets[0].cells['0,0'].raw).toBe('a')
    expect(data.sheets[0].cells['0,1'].raw).toBe('b')
    expect(data.sheets[0].cells['1,2'].raw).toBe('3')
  })

  it('handles quoted fields with embedded commas, quotes and newlines', () => {
    const csv = 'name,note\r\n"Smith, John","He said ""hi""\nnice"\r\n'
    const data = readCsv(csv)
    expect(data.sheets[0].cells['1,0'].raw).toBe('Smith, John')
    expect(data.sheets[0].cells['1,1'].raw).toBe('He said "hi"\nnice')
  })

  it('handles LF-only line endings and a trailing row with no newline', () => {
    const data = readCsv('x,y\n1,2\n3,4')
    expect(data.sheets[0].cells['2,0'].raw).toBe('3')
    expect(data.sheets[0].cells['2,1'].raw).toBe('4')
  })

  it('skips empty fields (no cell written)', () => {
    const data = readCsv('a,,c\n')
    expect(data.sheets[0].cells['0,0'].raw).toBe('a')
    expect(data.sheets[0].cells['0,1']).toBeUndefined()
    expect(data.sheets[0].cells['0,2'].raw).toBe('c')
  })
})

describe('writeCsv', () => {
  it('quotes fields that need it and uses CRLF', () => {
    const grid = [
      ['a', 'b,c', 'd"e'],
      ['1', '2', '3']
    ]
    const out = writeCsv(2, 3, (r, c) => grid[r][c])
    expect(out).toBe('a,"b,c","d""e"\r\n1,2,3\r\n')
  })

  it('round-trips through readCsv', () => {
    const grid = [
      ['name', 'quote'],
      ['Ann', 'She said "wow"']
    ]
    const csv = writeCsv(2, 2, (r, c) => grid[r][c])
    const back = readCsv(csv)
    expect(back.sheets[0].cells['1,1'].raw).toBe('She said "wow"')
  })
})

describe('CSV like Excel', () => {
  it('strips the BOM, reads 1,234 as a number and keeps dates formatted', async () => {
    const { store } = await import('../model/store')
    store.load(readCsv('﻿Code,Amt,Date,ISO\r\n007,"1,234",03.09.2026,2026-09-03\r\n'), 'c', null)
    const s = store.active
    expect(store.getRaw(s, 0, 0)).toBe('Code')
    expect(store.getValue(s, 1, 1)).toBe(1234)
    expect(store.getDisplay(s, 1, 3).text).toBe('2026-09-03')
    expect(store.getDisplay(s, 1, 2).text).not.toMatch(/^\d{5}$/)
  })
})

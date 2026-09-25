import { describe, it, expect } from 'vitest'
import { store } from '../model/store'

describe('Excel compatibility functions', () => {
  it('evaluates bare TRUE/FALSE, CONCAT and TEXTJOIN', () => {
    store.newWorkbook()
    const s = store.active
    store.setRangeRaw(s, 0, 0, [['10'], ['20'], [''], ['x']])
    const f = (raw: string): string => {
      store.setRangeRaw(s, 9, 0, [[raw]])
      return store.getDisplay(s, 9, 0).text
    }
    expect(f('=VLOOKUP(20,A1:A2,1,FALSE)')).toBe('20')
    expect(f('=IF(TRUE,"y","n")')).toBe('y')
    expect(f('=HLOOKUP(10,A1:A2,1,FALSE)')).toBe('10')
    expect(f('=CONCAT("a",A1,"b")')).toBe('a10b')
    expect(f('=CONCAT(A1:A4)')).toBe('1020x')
    expect(f('=TEXTJOIN(",",TRUE,A1:A4)')).toBe('10,20,x')
    expect(f('=TEXTJOIN("-",FALSE,A1:A4)')).toBe('10-20--x')
    expect(store.getRaw(s, 9, 0)).toBe('=TEXTJOIN("-",FALSE,A1:A4)')
  })
  it('keeps TRUE/FALSE working in a loaded workbook', () => {
    store.load({ sheets: [{ name: 'S', cells: { '0,0': { raw: '5' }, '1,0': { raw: '=IF(FALSE,1,A1)' } } }] } as any, 'x.xlsx', null)
    expect(store.getDisplay(store.active, 1, 0).text).toBe('5')
  })
})

describe('SUBTOTAL / AGGREGATE with filtered rows', () => {
  it('skips rows hidden by the filter and recalculates when the filter changes', () => {
    store.newWorkbook()
    const s = store.active
    store.setRangeRaw(s, 0, 0, [['h'], ['1'], ['2'], ['3'], ['4']])
    store.setRangeRaw(s, 6, 0, [['=SUBTOTAL(9,A2:A5)'], ['=SUBTOTAL(3,A2:A5)'], ['=AGGREGATE(9,5,A2:A5)'], ['=SUBTOTAL(109,A2:A5)'], ['=AGGREGATE(14,3,A2:A5,1)']])
    const v = (r: number) => store.getDisplay(s, r, 0).text
    expect([v(6), v(7), v(8)]).toEqual(['10', '4', '10'])
    store.updateLayout(s, 'Filter', (l) => {
      l.autoFilter = { range: { r1: 0, c1: 0, r2: 4, c2: 0 }, columns: {}, filteredOutRows: [2, 3] }
      l.hiddenRows = { 2: true, 3: true }
    })
    expect([v(6), v(7), v(8), v(10)]).toEqual(['5', '2', '5', '4'])
    store.updateLayout(s, 'Hide', (l) => { l.hiddenRows = { ...l.hiddenRows, 4: true } })
    expect([v(6), v(9)]).toEqual(['5', '1'])
    store.undo(); store.undo()
    expect(v(6)).toBe('10')
  })
})

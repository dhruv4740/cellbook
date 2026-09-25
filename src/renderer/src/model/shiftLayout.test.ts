import { describe, it, expect } from 'vitest'
import { shiftRange } from './shiftLayout'
import { store } from './store'

describe('shiftRange', () => {
  const rg = { r1: 2, c1: 0, r2: 5, c2: 0 }
  it('grows when rows are inserted inside, moves when inserted above', () => {
    expect(shiftRange(rg, 'r', 3, 2)).toEqual({ r1: 2, c1: 0, r2: 7, c2: 0 })
    expect(shiftRange(rg, 'r', 0, 1)).toEqual({ r1: 3, c1: 0, r2: 6, c2: 0 })
    expect(shiftRange(rg, 'r', 6, 1)).toEqual(rg)
  })
  it('shrinks, moves or disappears on delete', () => {
    expect(shiftRange(rg, 'r', 4, 3)).toBeTruthy()
    expect(shiftRange(rg, 'r', 4, -3)).toEqual({ r1: 2, c1: 0, r2: 3, c2: 0 })
    expect(shiftRange(rg, 'r', 0, -3)).toEqual({ r1: 0, c1: 0, r2: 2, c2: 0 })
    expect(shiftRange(rg, 'r', 1, -2)).toEqual({ r1: 1, c1: 0, r2: 3, c2: 0 })
    expect(shiftRange(rg, 'r', 2, -4)).toBeNull()
  })
})

describe('insert/delete rows keep layout attached to cells', () => {
  it('moves notes, CF, validation, filter and merges', () => {
    store.newWorkbook()
    const s = store.active
    store.setRangeRaw(s, 0, 0, [['h'], ['a'], ['b'], ['c']])
    store.updateLayout(s, 'setup', (l) => {
      l.notes = { '2,0': { text: 'on b', author: 'x' } }
      l.conditionalFormats = [{ id: 'c', priority: 1, type: 'cellIs', operator: 'greaterThan', value: '1', ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }] }]
      l.dataValidations = [{ id: 'd', allow: 'list', formula1: '"x,y"', ranges: [{ r1: 3, c1: 0, r2: 3, c2: 0 }] }] as any
      l.autoFilter = { range: { r1: 0, c1: 0, r2: 3, c2: 0 }, columns: {}, filteredOutRows: [3] }
      l.merges = [{ r1: 2, c1: 1, r2: 5, c2: 1 }]
    })
    store.insertRows(s, 1, 1)
    expect(store.getRaw(s, 3, 0)).toBe('b')
    expect(Object.keys(s.layout.notes)).toEqual(['3,0'])
    expect(s.layout.conditionalFormats![0].ranges[0]).toEqual({ r1: 2, c1: 0, r2: 4, c2: 0 })
    expect(s.layout.dataValidations![0].ranges[0]).toEqual({ r1: 4, c1: 0, r2: 4, c2: 0 })
    expect(s.layout.autoFilter!.range).toEqual({ r1: 0, c1: 0, r2: 4, c2: 0 })
    expect(s.layout.autoFilter!.filteredOutRows).toEqual([4])
    expect(s.layout.merges[0]).toEqual({ r1: 3, c1: 1, r2: 6, c2: 1 })
    store.deleteRows(s, 5, 3) // cuts the merge's bottom
    expect(s.layout.merges[0]).toEqual({ r1: 3, c1: 1, r2: 4, c2: 1 })
    store.undo(); store.undo()
    expect(Object.keys(s.layout.notes)).toEqual(['2,0'])
  })
})

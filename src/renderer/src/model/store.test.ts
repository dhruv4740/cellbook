import { describe, it, expect } from 'vitest'
import { WorkbookStore } from './store'

describe('store', () => {
  it('evaluates formulas and undoes', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '2')
    s.setRaw(sh, 1, 0, '3')
    s.setRaw(sh, 2, 0, '=SUM(A1:A2)')
    expect(s.getValue(sh, 2, 0)).toBe(5)
    s.undo()
    expect(s.getRaw(sh, 2, 0)).toBe('')
    s.redo()
    s.insertRows(sh, 0, 1)
    expect(s.getRaw(sh, 3, 0)).toBe('=SUM(A2:A3)')
    s.undo()
    expect(s.getRaw(sh, 2, 0)).toBe('=SUM(A1:A2)')
    s.clear(sh, [{ r1: 0, c1: 0, r2: 1, c2: 0 }])
    expect(s.getValue(sh, 2, 0)).toBe(0)
    s.undo()
    expect(s.getValue(sh, 2, 0)).toBe(5)
    s.addSheet()
    expect(s.sheets.map((x) => x.name)).toEqual(['Sheet1', 'Sheet2'])
    s.undo()
    expect(s.sheets.length).toBe(1)
  })
})

describe('input parsing', () => {
  it('parses US dates', () => {
    const s = new WorkbookStore()
    s.setRaw(s.active, 0, 0, '1/31/2025')
    expect(s.getValue(s.active, 0, 0)).toBe(45688)
  })
})

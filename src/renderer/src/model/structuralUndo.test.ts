import { describe, expect, it } from 'vitest'
import { WorkbookStore } from './store'

function make() {
  const s = new WorkbookStore()
  const sh = s.active
  ;[1, 2, 3, 4, 5].forEach((v, r) => s.setRaw(sh, r, 0, String(v)))
  s.setRaw(sh, 0, 1, '=SUM(A1:A5)')
  s.setRaw(sh, 1, 1, '=A3*2')
  s.setRaw(sh, 2, 1, '=A5+1')
  return { s, sh }
}

describe('insert/delete rows and columns with undo and redo', () => {
  it('delete rows: references shrink or become #REF!, undo restores everything', () => {
    const { s, sh } = make()
    s.deleteRows(sh, 2, 1) // A3
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A4)')
    expect(s.getValue(sh, 0, 1)).toBe(12)
    expect(s.getValue(sh, 1, 1)).toEqual({ error: '#REF!' })
    s.undo()
    expect(s.getRaw(sh, 2, 0)).toBe('3')
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A5)')
    expect(s.getRaw(sh, 1, 1)).toBe('=A3*2')
    expect(s.getValue(sh, 1, 1)).toBe(6)
    expect(s.getValue(sh, 2, 1)).toBe(6)
    s.redo()
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A4)')
    expect(s.getValue(sh, 1, 1)).toEqual({ error: '#REF!' })
    s.undo()
    expect(s.getValue(sh, 0, 1)).toBe(15)
  })
  it('insert rows: references shift, undo and redo are exact', () => {
    const { s, sh } = make()
    s.insertRows(sh, 1, 2)
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A7)')
    expect(s.getRaw(sh, 3, 1)).toBe('=A5*2')
    s.undo()
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A5)')
    expect(s.getRaw(sh, 1, 1)).toBe('=A3*2')
    s.redo()
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A7)')
  })
  it('delete and insert columns', () => {
    const { s, sh } = make()
    s.insertCols(sh, 0, 1)
    expect(s.getRaw(sh, 0, 2)).toBe('=SUM(B1:B5)')
    s.undo()
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A5)')
    s.deleteCols(sh, 0, 1)
    expect(s.getValue(sh, 0, 0)).toEqual({ error: '#REF!' })
    s.undo()
    expect(s.getRaw(sh, 0, 0)).toBe('1')
    expect(s.getRaw(sh, 0, 1)).toBe('=SUM(A1:A5)')
    expect(s.getValue(sh, 0, 1)).toBe(15)
  })
  it('keeps row formatting in step with the rows', () => {
    const { s, sh } = make()
    s.updateLayout(sh, 'h', (l) => { l.rowHeights[3] = 40 })
    s.insertRows(sh, 0, 1)
    expect(sh.layout.rowHeights[4]).toBe(40)
    s.undo()
    expect(sh.layout.rowHeights[3]).toBe(40)
    expect(sh.layout.rowHeights[4]).toBeUndefined()
  })
})

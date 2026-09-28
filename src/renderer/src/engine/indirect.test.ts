import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../model/store'

describe('INDIRECT', () => {
  it('resolves cell, range and cross-sheet text references and follows edits', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '10'); s.setRaw(sh, 1, 0, '20'); s.setRaw(sh, 2, 0, '30')
    s.setRaw(sh, 0, 2, '=INDIRECT("A2")')
    s.setRaw(sh, 1, 2, '=SUM(INDIRECT("A1:A3"))')
    s.setRaw(sh, 2, 2, '=INDIRECT("B"&1)')
    s.setRaw(sh, 3, 2, '=INDIRECT("Nope!A1")')
    expect(s.getValue(sh, 0, 2)).toBe(20)
    expect(s.getValue(sh, 1, 2)).toBe(60)
    s.setRaw(sh, 0, 1, '7')
    expect(s.getValue(sh, 2, 2)).toBe(7)
    s.setRaw(sh, 1, 0, '25')
    expect(s.getValue(sh, 0, 2)).toBe(25)
    expect(s.getValue(sh, 1, 2)).toBe(65)
    expect(s.getValue(sh, 3, 2)).toEqual({ error: '#REF!' })
  })
})

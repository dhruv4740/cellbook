import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import { cellErrorKind, nextErrorCell } from './errorCheck'

describe('error checking', () => {
  beforeEach(() => store.newWorkbook())
  it('flags numbers stored as text and formula errors, not ordinary text or numbers', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, "'123") // text that looks like a number
    store.setRaw(sh, 1, 0, '123') // a real number
    store.setRaw(sh, 2, 0, 'abc')
    store.setRaw(sh, 3, 0, '=1/0')
    store.setRaw(sh, 4, 0, "'1,234.5")
    expect(cellErrorKind(sh, 0, 0)).toBe('Number Stored as Text')
    expect(cellErrorKind(sh, 1, 0)).toBeNull()
    expect(cellErrorKind(sh, 2, 0)).toBeNull()
    expect(cellErrorKind(sh, 3, 0)).toMatch(/error/)
    expect(cellErrorKind(sh, 4, 0)).toBe('Number Stored as Text')
    store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } })
    expect(nextErrorCell(sh)).toMatchObject({ r: 3, c: 0 })
  })
})

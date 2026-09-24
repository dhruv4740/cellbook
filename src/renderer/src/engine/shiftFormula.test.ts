import { describe, it, expect } from 'vitest'
import { shiftFormula } from './shiftFormula'

describe('shiftFormula', () => {
  it('shifts a simple relative reference', () => {
    expect(shiftFormula('=A1', 1, 1)).toBe('=B2')
  })
  it('shifts negative (moving up/left)', () => {
    expect(shiftFormula('=C5', -2, -1)).toBe('=B3')
  })
  it('respects absolute column and row', () => {
    expect(shiftFormula('=$A$1', 5, 5)).toBe('=$A$1')
    expect(shiftFormula('=A$1', 5, 5)).toBe('=F$1')
    expect(shiftFormula('=$A1', 5, 5)).toBe('=$A6')
  })
  it('shifts ranges', () => {
    expect(shiftFormula('=SUM(A1:B2)', 1, 1)).toBe('=SUM(B2:C3)')
  })
  it('shifts sheet-qualified refs', () => {
    expect(shiftFormula('=Sheet2!A1', 1, 0)).toBe('=Sheet2!A2')
    expect(shiftFormula("='My Sheet'!A1", 0, 1)).toBe("='My Sheet'!B1")
  })
  it('leaves string literals untouched', () => {
    expect(shiftFormula('=CONCATENATE("A1",A1)', 1, 0)).toBe('=CONCATENATE("A1",A2)')
  })
  it('does not touch function names that look like refs', () => {
    expect(shiftFormula('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)')
  })
  it('yields #REF! when shifted off the top/left of the sheet', () => {
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!')
    expect(shiftFormula('=A1', 0, -1)).toBe('=#REF!')
  })
  it('leaves non-formula text alone', () => {
    expect(shiftFormula('hello A1', 1, 1)).toBe('hello A1')
  })
  it('handles multiple refs in one formula', () => {
    expect(shiftFormula('=A1+B2*C3', 1, 1)).toBe('=B2+C3*D4')
  })
})

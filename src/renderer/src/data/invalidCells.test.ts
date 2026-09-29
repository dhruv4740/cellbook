import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { addDataValidation, findInvalidCells } from './dataValidation'

describe('Circle Invalid Data', () => {
  it('finds values that break a whole-number rule', () => {
    const s = new WorkbookStore()
    const sh = s.active
    ;['5', '50', 'abc', '', '7'].forEach((v, r) => v && s.setRaw(sh, r, 0, v))
    addDataValidation(s, sh, { ranges: [{ r1: 0, c1: 0, r2: 4, c2: 0 }], allow: 'whole', operator: 'between', formula1: '1', formula2: '10', ignoreBlank: true })
    expect(findInvalidCells(s, sh).sort()).toEqual(['1,0', '2,0'])
  })
})

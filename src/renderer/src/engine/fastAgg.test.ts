import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../model/store'

// Semantics were checked against real Excel with scripts/excel-oracle/func_oracle.py (83 cases, 0 differences).
function sheet() {
  const s = new WorkbookStore()
  const sh = s.active
  const A = ['apple', 'Apple', 'banana', 'apricot', 'cherry', 'APPLE', '', '', 'a1', "'5", "'5", 'banana']
  A.forEach((v, r) => { if (v !== '') s.setRaw(sh, r, 0, v) })
  ;[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].forEach((v, r) => s.setRaw(sh, r, 1, String(v)))
  ;['10', '5', '5', '', '5', 'TRUE', 'abc', '0', '1', '5', '1', '7.5'].forEach((v, r) => { if (v !== '') s.setRaw(sh, r, 2, v) })
  ;[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].forEach((v, r) => s.setRaw(sh, r, 3, String(v)))
  let n = 0
  const val = (f: string) => { s.setRaw(sh, n, 10, '=' + f); return s.getValue(sh, n++, 10) }
  return val
}

describe('fast aggregate functions match Excel', () => {
  const v = sheet()
  it('COUNTIF text criteria', () => {
    expect(v('COUNTIF(A1:A12,"apple")')).toBe(3)
    expect(v('COUNTIF(A1:A12,"app*")')).toBe(3)
    expect(v('COUNTIF(A1:A12,"?pple")')).toBe(3)
    expect(v('COUNTIF(A1:A12,"<>apple")')).toBe(9)
    expect(v('COUNTIF(A1:A12,"5")')).toBe(2)
  })
  it('COUNTIF numeric criteria and mixed cells', () => {
    expect(v('COUNTIF(C1:C12,">4")')).toBe(6)
    expect(v('COUNTIF(C1:C12,5)')).toBe(4)
    expect(v('COUNTIF(C1:C12,"<>5")')).toBe(8)
  })
  it('SUMIF / SUMIFS / AVERAGEIF(S) / MAXIFS / MINIFS', () => {
    expect(v('SUMIF(A1:A12,"apple",B1:B12)')).toBe(9)
    expect(v('SUMIFS(B1:B12,A1:A12,"a*",C1:C12,">1")')).toBe(v('SUMIFS(B1:B12,C1:C12,">1",A1:A12,"a*")'))
    expect(v('AVERAGEIF(A1:A12,"apple",B1:B12)')).toBe(3)
    expect(v('AVERAGEIF(A1:A12,"nomatch",B1:B12)')).toEqual({ error: '#DIV/0!' })
    expect(v('MAXIFS(B1:B12,A1:A12,"apple")')).toBe(6)
    expect(v('MINIFS(B1:B12,A1:A12,"banana")')).toBe(3)
    expect(v('MAXIFS(B1:B12,A1:A12,"zzz")')).toBe(0)
  })
  it('XLOOKUP', () => {
    expect(v('XLOOKUP("banana",A1:A12,B1:B12)')).toBe(3)
    expect(v('XLOOKUP("BANANA",A1:A12,B1:B12,"none",0,-1)')).toBe(12)
    expect(v('XLOOKUP("zzz",A1:A12,B1:B12,"none")')).toBe('none')
    expect(v('XLOOKUP("zzz",A1:A12,B1:B12)')).toEqual({ error: '#N/A' })
    expect(v('XLOOKUP(6.5,D1:D12,B1:B12,"x",-1)')).toBe(6)
    expect(v('XLOOKUP("ap*",A1:A12,B1:B12,"x",2)')).toBe(1)
    expect(v('XLOOKUP("cherry",A1:A12,B1:C12)')).toBe(5)
  })
  it('INDEX', () => {
    expect(v('INDEX(B1:B12,3)')).toBe(3)
    expect(v('INDEX(B1:C12,2,2)')).toBe(5)
    expect(v('INDEX(C1:C12,4)')).toBe(0)
    expect(v('SUM(INDEX(B1:C12,0,2))')).toBe(v('SUM(C1:C12)'))
    expect(v('INDEX(B1:B12,13)')).toEqual({ error: '#REF!' })
  })
  it('stays correct after edits', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, 'x'); s.setRaw(sh, 0, 1, '5')
    s.setRaw(sh, 1, 0, 'x'); s.setRaw(sh, 1, 1, '7')
    s.setRaw(sh, 0, 3, '=SUMIF(A1:A2,"x",B1:B2)')
    s.setRaw(sh, 1, 3, '=XLOOKUP("y",A1:A2,B1:B2,"none")')
    expect(s.getValue(sh, 0, 3)).toBe(12)
    s.setRaw(sh, 1, 0, 'y')
    expect(s.getValue(sh, 0, 3)).toBe(5)
    expect(s.getValue(sh, 1, 3)).toBe(7)
    s.setRaw(sh, 0, 1, '100')
    expect(s.getValue(sh, 0, 3)).toBe(100)
    s.insertRows(sh, 0, 1)
    expect(s.getValue(sh, 1, 3)).toBe(100)
  })
})

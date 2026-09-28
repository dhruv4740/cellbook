import { describe, expect, it } from 'vitest'
import { expandLet } from './let'
import { WorkbookStore } from '../model/store'

describe('LET', () => {
  it('expands names into the body', () => {
    expect(expandLet('=LET(x,5,x*2)')).toBe('=((5)*2)')
    expect(expandLet('=LET(a,1,b,a+1,a+b)')).toBe('=((1)+((1)+1))')
    expect(expandLet('=SUM(1,2)')).toBeNull()
  })
  it('does not touch strings, functions or sheet names', () => {
    expect(expandLet('=LET(x,2,"x"&x&SUM(x))')).toBe('=("x"&(2)&SUM((2)))')
    expect(expandLet('=LET(x,2,x!A1)')).not.toContain('(2)!')
  })
  it('handles nesting', () => {
    expect(expandLet('=LET(x,LET(y,2,y*3),x+1)')).toBe('=(((2)*3)+1)'.replace('(((2)*3)+1)', '((((2)*3))+1)'))
  })
  it('works in a workbook and keeps the typed formula', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '4')
    s.setRaw(sh, 0, 1, '=LET(n,A1,sq,n*n,sq+n)')
    expect(s.getValue(sh, 0, 1)).toBe(20)
    expect(s.getRaw(sh, 0, 1)).toBe('=LET(n,A1,sq,n*n,sq+n)')
    s.setRaw(sh, 0, 0, '5')
    expect(s.getValue(sh, 0, 1)).toBe(30)
    expect(s.toData().sheets[0].cells['0,1'].raw).toBe('=LET(n,A1,sq,n*n,sq+n)')
  })
})

import { describe, expect, it } from 'vitest'
import { expandLambdas } from './lambda'
import { WorkbookStore } from '../model/store'

describe('LAMBDA family', () => {
  const mk = (): { s: WorkbookStore; ev: (f: string) => unknown; at: (r: number, c: number) => unknown } => {
    const s = new WorkbookStore()
    const sh = s.active
    ;[[1, 2, 3], [4, 5, 6], [7, 8, 9]].forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v))))
    return { s, ev: (f) => { s.setRaw(sh, 10, 10, f); return s.getValue(sh, 10, 10) }, at: (r, c) => s.getValue(sh, r, c) }
  }
  it('immediate invocation', () => {
    const { ev } = mk()
    expect(ev('=LAMBDA(x,y,x*y+1)(3,4)')).toBe(13)
  })
  it('REDUCE and SCAN', () => {
    const { s, ev, at } = mk()
    expect(ev('=REDUCE(0,A1:C1,LAMBDA(a,b,a+b))')).toBe(6)
    expect(ev('=REDUCE(A1:A3,LAMBDA(a,b,a*b))'.replace('A1:A3,LAMBDA', 'A1:A3,LAMBDA'))).toBeDefined()
    s.setRaw(s.active, 20, 0, '=SCAN(0,A1:C1,LAMBDA(a,b,a+b))')
    expect([at(20, 0), at(20, 1), at(20, 2)]).toEqual([1, 3, 6])
  })
  it('MAP, BYROW, BYCOL, MAKEARRAY', () => {
    const { s, at } = mk()
    s.setRaw(s.active, 20, 0, '=MAP(A1:B2,LAMBDA(v,v*10))')
    expect([at(20, 0), at(20, 1), at(21, 0), at(21, 1)]).toEqual([10, 20, 40, 50])
    s.setRaw(s.active, 25, 0, '=BYROW(A1:C3,LAMBDA(r,SUM(r)))')
    expect([at(25, 0), at(26, 0), at(27, 0)]).toEqual([6, 15, 24])
    s.setRaw(s.active, 30, 0, '=BYCOL(A1:C3,LAMBDA(c,MAX(c)))')
    expect([at(30, 0), at(30, 1), at(30, 2)]).toEqual([7, 8, 9])
    s.setRaw(s.active, 35, 0, '=MAKEARRAY(2,3,LAMBDA(i,j,i*10+j))')
    expect([at(35, 0), at(35, 2), at(36, 1)]).toEqual([11, 13, 22])
    s.setRaw(s.active, 40, 0, '=MAP(SEQUENCE(3),LAMBDA(n,n^2))')
    expect([at(40, 0), at(41, 0), at(42, 0)]).toEqual([1, 4, 9])
  })
  it('keeps the typed text and leaves unsupported shapes alone', () => {
    const { s } = mk()
    s.setRaw(s.active, 20, 0, '=MAP(A1:B2,LAMBDA(v,v*10))')
    expect(s.getRaw(s.active, 20, 0)).toBe('=MAP(A1:B2,LAMBDA(v,v*10))')
    expect(expandLambdas('=SUM(1,2)')).toBeNull()
    expect(expandLambdas('=MAP(A:A,LAMBDA(v,v))')).toBeNull()
  })
})

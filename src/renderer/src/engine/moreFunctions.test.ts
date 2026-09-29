import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'

const ev = (f: string): unknown => { const sh = store.active; store.setRaw(sh, 30, 10, f); return store.getValue(sh, 30, 10) }
const at = (r: number, c: number): unknown => store.getValue(store.active, r, c)

describe('more functions', () => {
  beforeEach(() => {
    store.newWorkbook()
    const sh = store.active
    ;[[3, 30], [1, 10], [2, 20], [5, 50]].forEach((row, r) => row.forEach((v, c) => store.setRaw(sh, r, c, String(v))))
  })
  it('scalars', () => {
    expect(ev('=NUMBERVALUE("1.234,5",",",".")')).toBe(1234.5)
    expect(ev('=NUMBERVALUE("50%")')).toBe(0.5)
    expect(ev('=TYPE("a")')).toBe(2)
    expect(ev('=TYPE(1)')).toBe(1)
    expect(ev('=FORECAST(6,B1:B4,A1:A4)')).toBeCloseTo(60, 6)
    expect(ev('=INTERCEPT(B1:B4,A1:A4)')).toBeCloseTo(0, 6)
    expect(ev('=MDETERM(MUNIT(3))')).toBe(1)
    expect(ev('=PERCENTRANK(A1:A4,2)')).toBeCloseTo(0.333, 3)
    expect(ev('=XMATCH(5,A1:A4)')).toBe(4)
    expect(ev('=XMATCH(4,A1:A4,-1)')).toBe(1)
    expect(ev('=XMATCH(9,A1:A4)')).toMatchObject({ error: '#N/A' })
  })
  it('array shaping spills', () => {
    const sh = store.active
    store.setRaw(sh, 10, 0, '=TAKE(A1:B4,2)')
    expect([at(10, 0), at(10, 1), at(11, 0), at(11, 1)]).toEqual([3, 30, 1, 10])
    store.setRaw(sh, 10, 4, '=DROP(A1:B4,3)')
    expect([at(10, 4), at(10, 5)]).toEqual([5, 50])
    store.setRaw(sh, 14, 0, '=CHOOSECOLS(A1:B4,2,1)')
    expect([at(14, 0), at(14, 1), at(17, 0), at(17, 1)]).toEqual([30, 3, 50, 5])
    store.setRaw(sh, 20, 0, '=SORTBY(A1:B4,A1:A4,1)')
    expect([at(20, 0), at(21, 0), at(22, 0), at(23, 0)]).toEqual([1, 2, 3, 5])
    expect([at(20, 1), at(23, 1)]).toEqual([10, 50])
    store.setRaw(sh, 26, 0, '=TOCOL(A1:B2)')
    expect([at(26, 0), at(27, 0), at(28, 0), at(29, 0)]).toEqual([3, 30, 1, 10])
    store.setRaw(sh, 26, 3, '=TOROW(A1:B2)')
    expect([at(26, 3), at(26, 4), at(26, 5), at(26, 6)]).toEqual([3, 30, 1, 10])
  })
  it('matrix and frequency', () => {
    const sh = store.active
    store.setRaw(sh, 10, 0, '=MUNIT(2)')
    expect([at(10, 0), at(10, 1), at(11, 0), at(11, 1)]).toEqual([1, 0, 0, 1])
    store.setRaw(sh, 14, 0, '2'); store.setRaw(sh, 14, 1, '0'); store.setRaw(sh, 15, 0, '0'); store.setRaw(sh, 15, 1, '4')
    store.setRaw(sh, 18, 0, '=MINVERSE(A15:B16)')
    expect([at(18, 0), at(19, 1)]).toEqual([0.5, 0.25])
    store.setRaw(sh, 22, 0, '=FREQUENCY(A1:A4,A1:A2)')
    expect([at(22, 0), at(23, 0), at(24, 0)]).toEqual([1, 2, 1]) // bins 1,3: <=1 -> 1, <=3 -> 2, above -> 1
  })
})

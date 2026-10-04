import { describe, it, expect } from 'vitest'
import { planTotals } from './quickAnalysis'

// Row 0 is a text header; B/C hold numbers, A holds labels.
const isNum = (r: number, c: number) => r > 0 && c > 0
const sel = { r1: 0, c1: 0, r2: 3, c2: 2 }

describe('planTotals', () => {
  it('writes column totals below the selection, skipping label columns', () => {
    expect(planTotals(sel, 'SUM', 'col', isNum)).toEqual([
      { r: 4, c: 1, raw: '=SUM(B1:B4)' },
      { r: 4, c: 2, raw: '=SUM(C1:C4)' }
    ])
  })
  it('writes row totals right of the selection, skipping the header row', () => {
    expect(planTotals(sel, 'AVERAGE', 'row', isNum).map((x) => [x.r, x.c, x.raw])).toEqual([
      [1, 3, '=AVERAGE(A2:C2)'], [2, 3, '=AVERAGE(A3:C3)'], [3, 3, '=AVERAGE(A4:C4)']
    ])
  })
  it('writes nothing when the selection has no numbers', () => {
    expect(planTotals(sel, 'COUNT', 'col', () => false)).toEqual([])
  })
})

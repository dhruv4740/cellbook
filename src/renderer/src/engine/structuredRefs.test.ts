import { describe, it, expect } from 'vitest'
import { resolveStructuredRefs, type TableDef } from './structuredRefs'

// Orders on sheet Dispatch at A1:F6, header row 1, data rows 2-6, like the table Excel built in the check.
const orders: TableDef = { name: 'Orders', sheet: 'Dispatch', r1: 0, c1: 0, r2: 5, c2: 5, headerRows: 1, totalsRows: 0, columns: ['Code', 'Party', 'Qty', 'Amount', 'Date', 'Unit Rate'] }
const r = (f: string, sheet = 'Dispatch', row = 7) => resolveStructuredRefs(f, [orders], sheet, row)

describe('structured references become ranges', () => {
  it('column, areas and ranges of columns', () => {
    expect(r('SUM(Orders[Qty])')).toBe('SUM($C$2:$C$6)')
    expect(r('COUNTA(Orders[[#All],[Code]])')).toBe('COUNTA($A$1:$A$6)')
    expect(r('ROWS(Orders[#Data])')).toBe('ROWS($A$2:$F$6)')
    expect(r('ROWS(Orders[])')).toBe('ROWS($A$2:$F$6)')
    expect(r('Orders[[#Headers],[Party]]')).toBe('$B$1')
    expect(r('SUM(Orders[[Qty]:[Amount]])')).toBe('SUM($C$2:$D$6)')
    expect(r('SUM(Orders[[Unit Rate]])')).toBe('SUM($F$2:$F$6)')
  })
  it('this row, from inside the table (row index 3 = Excel row 4)', () => {
    expect(r('Orders[[#This Row],[Qty]]*2', 'Dispatch', 3)).toBe('$C$4*2')
    expect(r('Orders[@Qty]*2', 'Dispatch', 3)).toBe('$C$4*2')
  })
  it('adds the sheet name from another sheet and leaves strings and unknown names alone', () => {
    expect(r('SUM(Orders[Qty])', 'Summary')).toBe('SUM(Dispatch!$C$2:$C$6)')
    expect(r('"Orders[Qty]"&Other[Qty]')).toBe('"Orders[Qty]"&Other[Qty]')
    expect(r('SUM(Orders[Nope])')).toBe('SUM(Orders[Nope])')
  })
})

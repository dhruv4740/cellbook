import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { renderPivot, refreshAllPivots, PivotDef } from './pivot'

function setup() {
  const s = new WorkbookStore()
  const sh = s.active
  const rows = [
    ['Source', 'Dest', 'Qty', 'Code'],
    ['Pune', 'Delhi', '5', 'A1'],
    ['Mumbai', 'Delhi', '7', 'A2'],
    ['Pune', 'Surat', '3', 'A3'],
    ['Mumbai', 'Surat', '10', 'A4'],
    ['Mumbai', 'Delhi', '1', ''],
    ['', 'Delhi', '2', 'A6']
  ]
  rows.forEach((r, i) => r.forEach((v, j) => s.setRaw(sh, i, j, v)))
  s.addSheet('Pivot')
  return { s, src: s.sheets[0], dst: s.sheets[1] }
}
const grid = (s: WorkbookStore, sh: ReturnType<typeof setup>['dst'], rows: number, cols: number) =>
  Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => s.getDisplay(sh, r, c).text))

describe('PivotTable', () => {
  it('sum by one row field with grand total, blanks last', () => {
    const { s, dst } = setup()
    const def: PivotDef = { id: 'p1', sourceSheet: 'Sheet1', source: { r1: 0, c1: 0, r2: 6, c2: 3 }, anchor: { r: 0, c: 0 }, rows: [0], cols: [], values: [{ col: 2, func: 'sum' }] }
    expect(renderPivot(s, dst, def)).toBeNull()
    expect(grid(s, dst, 5, 2)).toEqual([
      ['Row Labels', 'Sum of Qty'], ['Mumbai', '18'], ['Pune', '8'], ['(blank)', '2'], ['Grand Total', '28']
    ])
  })
  it('row + column fields, count, refresh after source edit', () => {
    const { s, src, dst } = setup()
    const def: PivotDef = { id: 'p2', sourceSheet: 'Sheet1', source: { r1: 0, c1: 0, r2: 6, c2: 3 }, anchor: { r: 2, c: 1 }, rows: [0], cols: [1], values: [{ col: 3, func: 'count' }] }
    renderPivot(s, dst, def)
    expect(grid(s, dst, 8, 5).slice(2).map((r) => r.slice(1))).toEqual([
      ['Count of Code', 'Column Labels', '', ''],
      ['Row Labels', 'Delhi', 'Surat', 'Grand Total'],
      ['Mumbai', '1', '1', '2'],
      ['Pune', '1', '1', '2'],
      ['(blank)', '1', '', '1'],
      ['Grand Total', '3', '2', '5']
    ])
    s.setRaw(src, 5, 3, 'A5')
    s.setRaw(src, 6, 0, 'Pune')
    refreshAllPivots(s)
    expect(grid(s, dst, 8, 5).slice(4).map((r) => r.slice(1))).toEqual([
      ['Mumbai', '2', '1', '3'],
      ['Pune', '2', '1', '3'],
      ['Grand Total', '4', '2', '6'],
      ['', '', '', ''] // the old (blank) row was cleared
    ])
  })
  it('two row fields nest with subtotals; undo removes the pivot', () => {
    const { s, dst } = setup()
    const def: PivotDef = { id: 'p3', sourceSheet: 'Sheet1', source: { r1: 0, c1: 0, r2: 4, c2: 3 }, anchor: { r: 0, c: 0 }, rows: [0, 1], cols: [], values: [{ col: 2, func: 'sum' }, { col: 2, func: 'average' }] }
    renderPivot(s, dst, def)
    expect(grid(s, dst, 8, 3)).toEqual([
      ['Row Labels', 'Sum of Qty', 'Average of Qty'],
      ['Mumbai', '17', '8.5'], ['Delhi', '7', '7'], ['Surat', '10', '10'],
      ['Pune', '8', '4'], ['Delhi', '5', '5'], ['Surat', '3', '3'],
      ['Grand Total', '25', '6.25']
    ])
    s.undo()
    expect(s.getDisplay(dst, 0, 0).text).toBe('')
    expect(dst.layout.pivots ?? []).toEqual([])
  })
})

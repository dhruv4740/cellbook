import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { applySubtotals, removeSubtotals } from './subtotal'

function setup() {
  const s = new WorkbookStore()
  const sh = s.active
  const rows = [['Source', 'Qty'], ['Mumbai', '5'], ['Mumbai', '7'], ['Pune', '3'], ['Surat', '1'], ['Surat', '2']]
  rows.forEach((r, i) => r.forEach((v, j) => s.setRaw(sh, i, j, v)))
  return { s, sh }
}

describe('Data > Subtotal', () => {
  it('adds a total after each group and a grand total (sum ignores inner subtotals)', () => {
    const { s, sh } = setup()
    expect(applySubtotals(s, sh, { r1: 0, c1: 0, r2: 5, c2: 1 }, { groupCol: 0, func: 0, addTo: [1], replace: true, summaryBelow: true })).toBeNull()
    const col = (c: number) => Array.from({ length: 10 }, (_, r) => s.getDisplay(sh, r, c).text)
    expect(col(0)).toEqual(['Source', 'Mumbai', 'Mumbai', 'Mumbai Total', 'Pune', 'Pune Total', 'Surat', 'Surat', 'Surat Total', 'Grand Total'])
    expect(col(1)).toEqual(['Qty', '5', '7', '12', '3', '3', '1', '2', '3', '18'])
    s.undo()
    expect(s.getDisplay(sh, 3, 0).text).toBe('Pune')
    expect(s.getDisplay(sh, 6, 0).text).toBe('')
  })
  it('replace re-applies and Remove All restores the data', () => {
    const { s, sh } = setup()
    applySubtotals(s, sh, { r1: 0, c1: 0, r2: 5, c2: 1 }, { groupCol: 0, func: 0, addTo: [1], replace: true, summaryBelow: true })
    applySubtotals(s, sh, { r1: 0, c1: 0, r2: 9, c2: 1 }, { groupCol: 0, func: 1, addTo: [1], replace: true, summaryBelow: true })
    expect(s.getDisplay(sh, 3, 0).text).toBe('Mumbai Count')
    expect(s.getDisplay(sh, 9, 1).text).toBe('5')
    removeSubtotals(s, sh, { r1: 0, c1: 0, r2: 9, c2: 1 })
    expect(Array.from({ length: 7 }, (_, r) => s.getDisplay(sh, r, 0).text)).toEqual(['Source', 'Mumbai', 'Mumbai', 'Pune', 'Surat', 'Surat', ''])
  })
})

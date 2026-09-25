// Random edit sequences: undoing everything must restore the workbook exactly, and redoing
// everything must reproduce the edited state exactly.
import { describe, it, expect } from 'vitest'
import { store } from './store'
import { sortRange } from '../data/sort'

function rng(seed: number): () => number {
  return () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 2 ** 32 }
}
const norm = (o: unknown): unknown => Array.isArray(o) ? o.map(norm) : o && typeof o === 'object' ? Object.fromEntries(Object.keys(o).sort().map((k) => [k, norm((o as Record<string, unknown>)[k])])) : o
const snap = () => JSON.stringify(norm({ d: store.toData(), n: store.sheets.map((s) => s.name) }))

describe('undo/redo fuzz', () => {
  for (const seed of [1, 7, 42, 99, 2026, 3, 5, 11, 13, 17, 23, 31]) {
    it(`seed ${seed}`, () => {
      const R = rng(seed)
      const int = (n: number) => Math.floor(R() * n)
      store.newWorkbook()
      const s = () => store.active
      const vals = ['1', '2.5', 'abc', '=A1+B2', '=SUM(A1:C5)', '', '=IF(A1>1,"x","y")', '10%', "'007", 'TRUE']
      const init: string[][] = []
      for (let r = 0; r < 8; r++) { init.push([]); for (let c = 0; c < 5; c++) init[r].push(vals[int(vals.length)]) }
      // Load (not edit) the starting data so the whole sequence below fits Excel's 100-step undo history.
      store.load({ sheets: [{ name: 'Sheet1', cells: Object.fromEntries(init.flatMap((row, r) => row.map((raw, c) => [`${r},${c}`, { raw }]))), layout: { notes: { '2,1': { text: 'n', author: 'a' } } } }] }, 'f', null)
      const start = snap()
      let ops = 0
      for (let i = 0; i < 95; i++) {
        const k = int(10)
        const r = int(10), c = int(6)
        if (k === 0) store.setRaw(s(), r, c, vals[int(vals.length)])
        else if (k === 1) store.setCells(s(), [{ r, c, raw: vals[int(vals.length)], style: R() < 0.5 ? { bold: true } : undefined }, { r: r + 1, c, raw: 'p' }], 'Paste')
        else if (k === 2) store.insertRows(s(), r, 1 + int(2))
        else if (k === 3) store.deleteRows(s(), r, 1 + int(2))
        else if (k === 4) store.insertCols(s(), c, 1)
        else if (k === 5) store.deleteCols(s(), c, 1)
        else if (k === 6) store.applyStyle(s(), [{ r1: r, c1: c, r2: r + 2, c2: c + 1 }], { fill: '#FF0000' })
        else if (k === 7) sortRange(store, s(), { r1: 0, c1: 0, r2: 9, c2: 4 }, { levels: [{ columnIndex: int(5), sortOn: 'values', order: R() < 0.5 ? 'asc' : 'desc' }], hasHeaders: false })
        else if (k === 8) store.setNote(s(), r, c, 'note ' + i)
        else store.updateLayout(s(), 'Merge', (l) => { l.merges = [...l.merges, { r1: r, c1: c, r2: r + 1, c2: c }] })
        ops++
      }
      const end = snap()
      let guard = 0
      while (store.canUndo && guard++ < 10000) store.undo()
      expect(snap()).toBe(start)
      guard = 0
      while (store.canRedo && guard++ < 10000) store.redo()
      expect(snap()).toBe(end)
      expect(ops).toBe(95)
    })
  }
})

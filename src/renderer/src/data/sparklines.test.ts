import { describe, it, expect } from 'vitest'
import { store } from '../model/store'
import { addSparklines, clearSparklines, sparkValues } from './sparklines'

describe('sparklines', () => {
  it('one per row, moves with inserted rows, undoes, clears', () => {
    const sh = store.active
    for (let r = 0; r < 3; r++) for (let c = 0; c < 4; c++) store.setRaw(sh, r, c, String(r * 10 + c))
    expect(addSparklines(sh, 'line', 'A1:D3', 'F1:F2')).toMatch(/doesn.t match/)
    expect(addSparklines(sh, 'column', 'A1:D3', 'F1:F3')).toBeNull()
    expect(sh.layout.sparklines).toHaveLength(3)
    expect(sparkValues(sh, sh.layout.sparklines![1])).toEqual([10, 11, 12, 13])
    store.undo()
    expect(sh.layout.sparklines ?? []).toHaveLength(0)
    store.redo()
    store.insertRows(sh, 0, 1)
    expect(sh.layout.sparklines!.map((s) => [s.r, s.src.r1])).toEqual([[1, 1], [2, 2], [3, 3]])
    clearSparklines(sh, [{ r1: 1, c1: 5, r2: 1, c2: 5 }])
    expect(sh.layout.sparklines).toHaveLength(2)
  })
})

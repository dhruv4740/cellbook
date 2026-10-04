import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { planPicture } from './copyPicture'

describe('planPicture', () => {
  it('lays out visible cells with their display text and skips hidden rows / columns', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, 'a'); s.setRaw(sh, 0, 1, '1234.5'); s.setRaw(sh, 1, 0, 'b'); s.setRaw(sh, 2, 0, 'c')
    sh.layout.colWidths[1] = 100
    sh.layout.hiddenRows[1] = true
    const plan = planPicture(sh, { r1: 0, c1: 0, r2: 2, c2: 1 }, true, s)!
    expect(plan.width).toBe(64 + 100)
    expect(plan.height).toBe(40)
    expect(plan.cells.map((c) => c.text)).toEqual(['a', '1234.5', 'c', ''])
    expect(plan.cells[1]).toMatchObject({ x: 64, y: 0, w: 100, align: 'right' })
    expect(plan.cells[2].y).toBe(20)
  })
  it('refuses oversized ranges', () => {
    const s = new WorkbookStore()
    expect(planPicture(s.active, { r1: 0, c1: 0, r2: 99999, c2: 50 }, false, s)).toBeNull()
  })
})

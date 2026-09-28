import { describe, it, expect } from 'vitest'
import { WorkbookStore } from './store'

describe('defined names (Name Manager)', () => {
  it('defines, edits, deletes with undo', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '10'); s.setRaw(sh, 1, 0, '20')
    expect(s.defineName('Qty', '=Sheet1!$A$1:$A$2')).toBeNull()
    s.setRaw(sh, 0, 1, '=SUM(Qty)')
    expect(s.getValue(sh, 0, 1)).toBe(30)
    expect(s.defineName('Qty', '=Sheet1!$A$1')).toMatch(/already exists/)
    expect(s.defineName('A1', '=1')).toMatch(/not valid/)
    expect(s.defineName('my name', '=1')).toMatch(/not valid/)
    expect(s.defineName('Qty', '=Sheet1!$A$1', undefined, { name: 'Qty' })).toBeNull()
    expect(s.getValue(sh, 0, 1)).toBe(10)
    s.undo()
    expect(s.getValue(sh, 0, 1)).toBe(30)
    s.deleteName('Qty')
    expect(s.names.length).toBe(0)
    expect(s.getDisplay(sh, 0, 1).text).toBe('#NAME?')
    s.undo()
    expect(s.getValue(sh, 0, 1)).toBe(30)
    expect(s.toData().names?.[0]).toMatchObject({ name: 'Qty', ref: 'Sheet1!$A$1:$A$2' })
  })
})

describe('hide / copy sheets', () => {
  it('keeps one sheet visible and survives save', () => {
    const s = new WorkbookStore()
    s.addSheet('Data')
    expect(s.hideSheet(1)).toBeNull()
    expect(s.activeIndex).toBe(0)
    expect(s.hideSheet(0)).toMatch(/at least one visible/)
    expect(s.toData().sheets[1].layout?.sheetHidden).toBe('hidden')
    s.unhideSheet(1)
    expect(s.sheets[1].layout.sheetHidden).toBeUndefined()
    s.undo()
    expect(s.sheets[1].layout.sheetHidden).toBe('hidden')
  })
  it('copies cells, formulas and styles; undo removes the copy', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '5'); s.setRaw(sh, 0, 1, '=A1*2')
    s.applyStyle(sh, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { bold: true })
    s.copySheet(0, 1)
    expect(s.sheets.map((x) => x.name)).toEqual(['Sheet1', 'Sheet1 (2)'])
    const c = s.sheets[1]
    expect(s.getValue(c, 0, 1)).toBe(10)
    expect(s.getStyle(c, 0, 0)?.bold).toBe(true)
    s.setRaw(c, 0, 0, '7')
    expect(s.getValue(sh, 0, 1)).toBe(10)
    s.undo(); s.undo()
    expect(s.sheets.length).toBe(1)
  })
})

describe('Calculate Now (F9)', () => {
  it('recomputes RAND and keeps names, formulas and undo', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '4')
    s.defineName('X', '=Sheet1!$A$1')
    s.setRaw(sh, 1, 0, '=X*2')
    s.setRaw(sh, 2, 0, '=RAND()')
    const seen = new Set<number>()
    for (let i = 0; i < 5; i++) { s.recalculate(); seen.add(s.getValue(sh, 2, 0) as number) }
    expect(seen.size).toBeGreaterThan(1)
    expect(s.getValue(sh, 1, 0)).toBe(8)
    s.setRaw(sh, 0, 0, '5')
    expect(s.getValue(sh, 1, 0)).toBe(10)
    s.undo()
    expect(s.getValue(sh, 1, 0)).toBe(8)
  })
})

describe('manual calculation mode', () => {
  it('keeps formula results stale until F9', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '1')
    s.setRaw(sh, 0, 1, '=A1*10')
    s.setCalcMode('manual')
    s.setRaw(sh, 0, 0, '5')
    expect(s.getValue(sh, 0, 0)).toBe(5)
    expect(s.getValue(sh, 0, 1)).toBe(10)
    expect(s.calcPending).toBe(true)
    s.setRaw(sh, 1, 1, '=A1+1') // a newly entered formula calculates
    expect(s.getValue(sh, 1, 1)).toBe(6)
    s.recalculate()
    expect(s.getValue(sh, 0, 1)).toBe(50)
    expect(s.calcPending).toBe(false)
    s.setCalcMode('auto')
    s.setRaw(sh, 0, 0, '2')
    expect(s.getValue(sh, 0, 1)).toBe(20)
  })
})

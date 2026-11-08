import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import { computeDataTable, saveScenario, showScenario, scenarioSummary, scenariosOf } from './whatIf'

describe('data table', () => {
  beforeEach(() => store.newWorkbook())
  it('two-variable table fills every combination and leaves the input cells alone', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, '2') // A1 = x
    store.setRaw(sh, 1, 0, '3') // A2 = y
    store.setRaw(sh, 2, 0, '=A1*10+A2') // A3 = formula
    // table at C1:F4 -> top-left C1 = formula, row values D1:F1 (x), column values C2:C4 (y)
    store.setRaw(sh, 0, 2, '=A3')
    ;[1, 2, 3].forEach((x, j) => store.setRaw(sh, 0, 3 + j, String(x)))
    ;[7, 8, 9].forEach((y, i) => store.setRaw(sh, 1 + i, 2, String(y)))
    computeDataTable(sh, { r1: 0, c1: 2, r2: 3, c2: 5 }, { r: 0, c: 0 }, { r: 1, c: 0 })
    expect(store.getRaw(sh, 1, 3)).toBe('17') // x=1, y=7
    expect(store.getRaw(sh, 3, 5)).toBe('39') // x=3, y=9
    expect(store.getRaw(sh, 0, 0)).toBe('2')
    expect(store.getRaw(sh, 1, 0)).toBe('3')
    store.undo()
    expect(store.getRaw(sh, 1, 3)).toBe('')
    expect(store.getRaw(sh, 0, 0)).toBe('2')
  })
  it('one-variable table (column input) evaluates each formula across the top row', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, '5')
    store.setRaw(sh, 0, 3, '=A1*2')
    store.setRaw(sh, 0, 4, '=A1+1')
    ;[10, 20].forEach((v, i) => store.setRaw(sh, 1 + i, 2, String(v)))
    computeDataTable(sh, { r1: 0, c1: 2, r2: 2, c2: 4 }, null, { r: 0, c: 0 })
    expect([store.getRaw(sh, 1, 3), store.getRaw(sh, 1, 4), store.getRaw(sh, 2, 3), store.getRaw(sh, 2, 4)]).toEqual(['20', '11', '40', '21'])
    expect(store.getRaw(sh, 0, 0)).toBe('5')
  })
})

describe('scenarios', () => {
  beforeEach(() => store.newWorkbook())
  it('show applies values (undoable) and summary builds a sheet with result cells', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, '100'); store.setRaw(sh, 1, 0, '5'); store.setRaw(sh, 2, 0, '=A1*A2')
    saveScenario(sh, { name: 'Low', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['50', '2'] })
    saveScenario(sh, { name: 'High', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['200', '9'] })
    expect(scenariosOf(sh).map((s) => s.name).sort()).toEqual(['High', 'Low'])
    showScenario(sh, scenariosOf(sh).find((s) => s.name === 'Low')!)
    expect(store.getValue(sh, 2, 0)).toBe(100)
    store.undo()
    expect(store.getValue(sh, 2, 0)).toBe(500)
    scenarioSummary(sh, [{ r: 2, c: 0 }])
    const out = store.active
    expect(out.name).toBe('Scenario Summary')
    const dump = Array.from({ length: 9 }, (_, r) => Array.from({ length: 7 }, (_c, c) => store.getRaw(out, r, c)).join('|')).join('\n')
    expect(dump).toContain('Low')
    expect(dump).toContain('1800') // High: 200 * 9
    expect(store.getRaw(sh, 0, 0)).toBe('100') // original cells restored
  })
})

import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import { applyStep, nextStep } from './evaluateFormula'

const run = (expr: string): string[] => {
  const seen = [expr]
  for (let i = 0; i < 20; i++) {
    const s = nextStep(expr, store.active)
    if (!s) break
    expr = applyStep(expr, s)
    seen.push(expr)
  }
  return seen
}

describe('evaluate formula', () => {
  beforeEach(() => store.newWorkbook())
  it('walks references, then innermost calls, down to a value', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, '4'); store.setRaw(sh, 1, 0, '6'); store.setRaw(sh, 2, 0, 'x')
    expect(run('IF(SUM(A1:A2)>5,A3&"!","no")')).toEqual([
      'IF(SUM(A1:A2)>5,A3&"!","no")',
      'IF(SUM({4;6})>5,A3&"!","no")',
      'IF(SUM({4;6})>5,"x"&"!","no")',
      'IF(10>5,"x"&"!","no")',
      '"x!"'
    ])
  })
  it('leaves quoted text that looks like a reference alone', () => {
    store.setRaw(store.active, 0, 0, '9')
    expect(run('"A1"&A1').slice(-1)).toEqual(['"A19"'])
  })
})

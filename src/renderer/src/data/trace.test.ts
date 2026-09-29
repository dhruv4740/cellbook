import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import { traceArrows } from './trace'

describe('trace arrows', () => {
  beforeEach(() => store.newWorkbook())
  it('walks precedents and dependents level by level', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, '1') // A1
    store.setRaw(sh, 0, 1, '=A1*2') // B1
    store.setRaw(sh, 0, 2, '=SUM(A1:B1)') // C1
    const one = traceArrows(0, 2, 'precedents', 1)
    expect(one.map((a) => [a.from.r1, a.from.c1, a.from.c2])).toEqual([[0, 0, 1]])
    const two = traceArrows(0, 2, 'precedents', 2)
    expect(two.length).toBe(2) // A1:B1 -> C1, then A1 -> B1
    const deps = traceArrows(0, 0, 'dependents', 1)
    expect(deps.map((a) => a.to.c1).sort()).toEqual([1, 2])
  })
})

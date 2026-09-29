import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import { applyNames, applicableNames } from './applyNames'

describe('apply names', () => {
  beforeEach(() => store.newWorkbook())
  it('replaces exact references with the defined names, leaves text and other refs alone', () => {
    const sh = store.active
    store.setRaw(sh, 0, 0, '10'); store.setRaw(sh, 1, 0, '20'); store.setRaw(sh, 2, 0, '5')
    expect(store.defineName('Rate', '=Sheet1!$A$3')).toBeNull()
    expect(store.defineName('Qty', '=Sheet1!$A$1:$A$2')).toBeNull()
    store.setRaw(sh, 0, 2, '=SUM(A1:A2)*A3')
    store.setRaw(sh, 1, 2, '="A3"&A1')
    expect(applicableNames().map((n) => n.name).sort()).toEqual(['Qty', 'Rate'])
    expect(applyNames(sh, ['Rate', 'Qty'])).toBe(1)
    expect(store.getRaw(sh, 0, 2)).toBe('=SUM(Qty)*Rate')
    expect(store.getValue(sh, 0, 2)).toBe(150)
    expect(store.getRaw(sh, 1, 2)).toBe('="A3"&A1')
    store.undo()
    expect(store.getRaw(sh, 0, 2)).toBe('=SUM(A1:A2)*A3')
  })
})

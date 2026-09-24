import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import * as cmds from './commands'

const sel = (r1: number, c1: number, r2 = r1, c2 = c1) =>
  store.setSelection({ ranges: [{ r1, c1, r2, c2 }], active: { r: r1, c: c1 } })
const raw = (r: number, c: number) => store.getRaw(store.active, r, c)

describe('commands', () => {
  beforeEach(() => store.newWorkbook())

  it('delete shift-left keeps cells before the deleted range', () => {
    store.setRangeRaw(store.active, 0, 0, [['a', 'b', 'c', 'd']])
    sel(0, 1)
    cmds.deleteShift('left')
    expect([raw(0, 0), raw(0, 1), raw(0, 2), raw(0, 3)]).toEqual(['a', 'c', 'd', ''])
    store.undo()
    expect([raw(0, 0), raw(0, 1), raw(0, 2), raw(0, 3)]).toEqual(['a', 'b', 'c', 'd'])
  })

  it('delete shift-up moves cells below into the gap', () => {
    store.setRangeRaw(store.active, 0, 0, [['1'], ['2'], ['3']])
    sel(1, 0)
    cmds.deleteShift('up')
    expect([raw(0, 0), raw(1, 0), raw(2, 0)]).toEqual(['1', '3', ''])
  })

  it('cut and paste onto the same cell keeps the value', async () => {
    store.setRaw(store.active, 0, 0, 'x')
    sel(0, 0)
    await cmds.cut()
    await cmds.paste()
    expect(raw(0, 0)).toBe('x')
  })

  it('cut overlapping move keeps the overlap', async () => {
    store.setRangeRaw(store.active, 0, 0, [['1'], ['2']])
    sel(0, 0, 1, 0)
    await cmds.cut()
    sel(1, 0)
    await cmds.paste()
    expect([raw(0, 0), raw(1, 0), raw(2, 0)]).toEqual(['', '1', '2'])
  })

  it('pastes from another sheet', async () => {
    store.setRaw(store.active, 0, 0, 'src')
    sel(0, 0)
    await cmds.copy()
    store.addSheet()
    sel(2, 2)
    await cmds.paste()
    expect(raw(2, 2)).toBe('src')
    await cmds.paste('link')
    expect(raw(2, 2)).toBe('=Sheet1!A1')
  })

  it('sort shifts relative references with their rows', () => {
    store.setRangeRaw(store.active, 0, 0, [['3', '=A1*10'], ['1', '=A2*10']])
    sel(0, 0, 1, 1)
    cmds.sortSelection('asc')
    expect([raw(0, 1), raw(1, 1)]).toEqual(['=A1*10', '=A2*10'])
    expect(store.getValue(store.active, 0, 1)).toBe(10)
  })
})

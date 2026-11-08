// Operations that restructure data must keep every piece attached to the right cell.
import { describe, it, expect } from 'vitest'
import { store } from './store'
import * as cmds from '../grid/commands'
import { readXlsx, writeXlsx } from '../io/xlsx'

function setup(): void {
  store.load({ sheets: [
    { name: 'Data', cells: { '0,0': { raw: 'Name' }, '0,1': { raw: 'Qty' }, '1,0': { raw: 'b' }, '1,1': { raw: '2' }, '2,0': { raw: 'c' }, '2,1': { raw: '3' }, '3,0': { raw: 'a' }, '3,1': { raw: '1', style: { fill: '#FFFF00' } }, '1,2': { raw: '=B2*10' } },
      layout: { notes: { '3,0': { text: 'note on a', author: 'x' } } } },
    { name: 'Summary', cells: { '0,0': { raw: '=SUM(Qtys)' }, '1,0': { raw: '=Data!A2' }, '2,0': { raw: '=Other!A1' } } },
    { name: 'Other', cells: { '0,0': { raw: '7' } } }
  ], names: [{ name: 'Qtys', ref: 'Data!$B$2:$B$4' }] }, 't', null)
}

describe('data integrity', () => {
  it('sort moves values, formats, formulas and notes together; undo restores', () => {
    setup()
    const d = store.sheets[0]
    store.setActiveSheet(0)
    store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 3, c2: 2 }], active: { r: 1, c: 0 } })
    cmds.sortSelection('asc')
    expect([1, 2, 3].map((r) => store.getRaw(d, r, 0))).toEqual(['a', 'b', 'c'])
    expect(store.getStyle(d, 1, 1)?.fill).toBe('#FFFF00')
    expect(store.getRaw(d, 2, 2)).toBe('=B3*10')
    expect(Object.keys(d.layout.notes)).toEqual(['1,0'])
    store.undo()
    expect([1, 2, 3].map((r) => store.getRaw(d, r, 0))).toEqual(['b', 'c', 'a'])
    expect(Object.keys(d.layout.notes)).toEqual(['3,0'])
  })
  it('renaming a sheet updates formulas and named ranges in the saved file', async () => {
    setup()
    store.renameSheet(0, 'Dispatch Data')
    expect(store.getRaw(store.sheets[1], 1, 0)).toBe("='Dispatch Data'!A2")
    expect(store.getDisplay(store.sheets[1], 0, 0).text).toBe('6')
    const back = await readXlsx(await writeXlsx(store.toData()))
    expect(back.names).toEqual([{ name: 'Qtys', ref: "'Dispatch Data'!$B$2:$B$4" }])
  })
  it('deleting a sheet turns references into #REF! and undo brings them back', () => {
    setup()
    store.deleteSheet(2)
    expect(store.getDisplay(store.sheets[1], 2, 0).text).toBe('#REF!')
    store.undo()
    expect(store.getDisplay(store.sheets[1], 2, 0).text).toBe('7')
  })
  it('named ranges grow when rows are inserted inside them', () => {
    setup()
    store.insertRows(store.sheets[0], 2, 1)
    store.setRaw(store.sheets[0], 2, 1, '10')
    expect(store.getDisplay(store.sheets[1], 0, 0).text).toBe('16')
    expect(store.toData().names).toEqual([{ name: 'Qtys', ref: 'Data!$B$2:$B$5' }])
  })
})

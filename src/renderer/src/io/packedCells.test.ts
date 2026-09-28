import { describe, expect, it } from 'vitest'
import { packCells } from './packedCells'
import { WorkbookStore, WorkbookData } from '../model/store'

describe('packed cells', () => {
  const cells: WorkbookData['sheets'][number]['cells'] = {
    '0,0': { raw: 'Name', style: { bold: true } },
    '0,1': { raw: '12.5', style: { numFmt: '0.00' } },
    '1,0': { raw: "'007" },
    '1,1': { raw: '=B1*2' },
    '3,2': { style: { fill: '#FFFF00' } },
    '4,0': { raw: 'ünïcödé ✓' }
  }
  it('loads exactly like the plain cell map', () => {
    const plain = new WorkbookStore()
    plain.load({ sheets: [{ name: 'S', cells }] }, 'a', null)
    const packed = new WorkbookStore()
    const p = packCells(cells)
    packed.load({ sheets: [{ name: 'S', cells: {}, packed: p }] }, 'a', null)
    expect(packed.toData().sheets[0].cells).toEqual(plain.toData().sheets[0].cells)
    expect(packed.getValue(packed.active, 1, 1)).toBe(25)
    expect(packed.getStyle(packed.active, 3, 2)?.fill).toBe('#FFFF00')
    expect(p.styles.length).toBe(3)
    expect(p.len[cellsIndex('3,2')]).toBe(-1)
  })
})

function cellsIndex(key: string): number {
  return ['0,0', '0,1', '1,0', '1,1', '3,2', '4,0'].indexOf(key)
}

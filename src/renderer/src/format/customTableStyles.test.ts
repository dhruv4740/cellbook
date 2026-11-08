import { describe, expect, it } from 'vitest'
import { TableStyleDef, upsertStyle } from './customTableStyles'
import { store } from '../model/store'
import { formatAsTable } from '../grid/commands'

const mk = (name: string, header = '#112233'): TableStyleDef => ({ name, header, band: '#ddeeff', line: '#445566', headerFont: '#ffffff' })

describe('custom table styles', () => {
  it('upsert adds new names and replaces the same name case-insensitively', () => {
    const a = upsertStyle([], mk('Mine'))
    expect(upsertStyle(a, mk('Other'))).toHaveLength(2)
    const r = upsertStyle(a, mk('mine', '#000000'))
    expect(r).toHaveLength(1)
    expect(r[0].header).toBe('#000000')
  })
  it('formatAsTable paints a custom style: header fill, banding, borders', () => {
    const sh = store.active
    for (let r = 0; r < 4; r++) { store.setRaw(sh, r, 0, r === 0 ? 'H' : String(r)); store.setRaw(sh, r, 1, String(r)) }
    store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 3, c2: 1 }], active: { r: 0, c: 0 } })
    formatAsTable(mk('Mine'))
    expect(store.getStyle(sh, 0, 0)).toMatchObject({ fill: '#112233', color: '#ffffff', bold: true })
    expect(store.getStyle(sh, 1, 0)!.fill).toBe('#ddeeff')
    expect(store.getStyle(sh, 2, 0)!.fill).toBeUndefined()
    expect(store.getStyle(sh, 1, 1)!.borders?.top?.color).toBe('#445566')
  })
})

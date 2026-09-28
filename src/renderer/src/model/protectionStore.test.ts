import { describe, expect, it } from 'vitest'
import { WorkbookStore } from './store'
import { defaultAllow } from './protection'
import { readXlsx, writeXlsx } from '../io/xlsx'

function protectedStore() {
  const s = new WorkbookStore()
  const sh = s.active
  for (let r = 0; r < 4; r++) { s.setRaw(sh, r, 0, String(r + 1)); s.setRaw(sh, r, 1, String((r + 1) * 10)) }
  s.applyStyle(sh, [{ r1: 1, c1: 1, r2: 1, c2: 1 }], { locked: false }) // B2 unlocked
  const msgs: string[] = []
  s.onProtectedEdit = (m) => msgs.push(m)
  s.setSheetProtection(sh, { allow: defaultAllow() })
  return { s, sh, msgs }
}

describe('sheet protection enforcement', () => {
  it('blocks edits to locked cells and allows unlocked ones', () => {
    const { s, sh, msgs } = protectedStore()
    s.setRaw(sh, 0, 0, '99')
    expect(s.getRaw(sh, 0, 0)).toBe('1')
    expect(msgs[0]).toMatch(/protected sheet/)
    s.setRaw(sh, 1, 1, '77')
    expect(s.getRaw(sh, 1, 1)).toBe('77')
  })
  it('blocks a paste/fill (setCells) if any target is locked, and clear', () => {
    const { s, sh } = protectedStore()
    s.setCells(sh, [{ r: 1, c: 1, raw: '5' }, { r: 2, c: 1, raw: '6' }], 'Paste')
    expect(s.getRaw(sh, 1, 1)).toBe('20')
    s.clear(sh, [{ r1: 0, c1: 0, r2: 3, c2: 1 }])
    expect(s.getRaw(sh, 0, 0)).toBe('1')
  })
  it('format / structure / filter permissions follow the allow list', () => {
    const { s, sh } = protectedStore()
    s.applyStyle(sh, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { bold: true })
    expect(s.getStyle(sh, 0, 0)?.bold).toBeUndefined()
    s.setColWidth(sh, [0], 120)
    expect(sh.layout.colWidths[0]).toBeUndefined()
    s.insertRows(sh, 1, 1)
    expect(s.getRaw(sh, 1, 0)).toBe('2')
    s.setSheetProtection(sh, { allow: { ...defaultAllow(), formatCells: true, formatColumns: true, insertRows: true, deleteRows: true } })
    s.applyStyle(sh, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { bold: true })
    expect(s.getStyle(sh, 0, 0)?.bold).toBe(true)
    s.setColWidth(sh, [0], 120)
    expect(sh.layout.colWidths[0]).toBe(120)
    s.insertRows(sh, 1, 1)
    expect(s.getRaw(sh, 2, 0)).toBe('2')
    s.deleteRows(sh, 2, 1) // row holds locked cells: refused even though "delete rows" is allowed
    expect(s.getRaw(sh, 2, 0)).toBe('2')
  })
  it('Undo never removes protection', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setColWidth(sh, [0], 100)
    s.setSheetProtection(sh, { allow: defaultAllow() })
    s.undo()
    expect(sh.layout.protection).toBeDefined()
  })
  it('locked structure blocks sheet changes', () => {
    const s = new WorkbookStore()
    s.setBookProtection({ lockStructure: true, lockWindows: false })
    s.addSheet()
    expect(s.sheets.length).toBe(1)
    expect(s.renameSheet(0, 'X')).toBeNull()
    expect(s.sheets[0].name).toBe('Sheet1')
    s.setBookProtection(undefined)
    s.addSheet()
    expect(s.sheets.length).toBe(2)
  })
  it('unlocked/locked state, protection and calc mode survive a save', async () => {
    const { s, sh } = protectedStore()
    s.setBookProtection({ lockStructure: true, lockWindows: false })
    s.calcMode = 'manual'
    s.setSheetProtection(sh, { allow: { ...defaultAllow(), sort: true, formatCells: true, selectLockedCells: false } })
    const back = await readXlsx(await writeXlsx(s.toData()))
    expect(back.sheets[0].layout?.protection?.allow).toEqual({ ...defaultAllow(), sort: true, formatCells: true, selectLockedCells: false })
    expect(back.sheets[0].cells['1,1'].style?.locked).toBe(false)
    expect(back.protection).toEqual({ lockStructure: true, lockWindows: false })
    expect(back.calcMode).toBe('manual')
  })
})

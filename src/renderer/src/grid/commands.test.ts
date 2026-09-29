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

  it('justify re-flows text to the selection width and undoes', () => {
    store.setRangeRaw(store.active, 0, 0, [['the quick brown fox'], ['jumps over the lazy dog']])
    sel(0, 0, 1, 0)
    const chars = Math.floor((store.colWidth(store.active, 0) / 7) * 0.95)
    cmds.justifySelection()
    const lines = Array.from({ length: 12 }, (_, r) => raw(r, 0)).filter(Boolean)
    expect(lines.join(' ')).toBe('the quick brown fox jumps over the lazy dog')
    for (const l of lines) expect(l.length).toBeLessThanOrEqual(Math.max(chars, 9))
    store.undo()
    expect([raw(0, 0), raw(1, 0)]).toEqual(['the quick brown fox', 'jumps over the lazy dog'])
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

  // ---- Ctrl+[ / Ctrl+]: precedents / dependents ----
  it('selectPrecedents selects the direct precedent of a formula cell', () => {
    store.setRaw(store.active, 0, 0, '5') // A1
    store.setRaw(store.active, 1, 0, '=A1*2') // A2
    sel(1, 0)
    cmds.selectPrecedents(false)
    const s = store.active.selection
    expect(s.active).toEqual({ r: 0, c: 0 })
    expect(s.ranges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 0 }])
  })

  it('selectDependents selects the direct dependent of a cell', () => {
    store.setRaw(store.active, 0, 0, '5') // A1
    store.setRaw(store.active, 1, 0, '=A1*2') // A2
    sel(0, 0)
    cmds.selectDependents(false)
    expect(store.active.selection.active).toEqual({ r: 1, c: 0 })
  })

  it('selectPrecedents(true) walks all levels of the chain', () => {
    store.setRaw(store.active, 0, 0, '1') // A1
    store.setRaw(store.active, 1, 0, '=A1+1') // A2
    store.setRaw(store.active, 2, 0, '=A2+1') // A3
    sel(2, 0)
    cmds.selectPrecedents(true)
    const cells = store.active.selection.ranges.map((r) => `${r.r1},${r.c1}`).sort()
    expect(cells).toEqual(['0,0', '1,0'])
  })

  it('selectPrecedents does nothing for a cell with no formula precedents', () => {
    store.setRaw(store.active, 0, 0, 'plain text')
    sel(0, 0)
    const before = store.active.selection
    cmds.selectPrecedents(false)
    expect(store.active.selection).toBe(before)
  })

  // ---- Ctrl+/: current array ----
  it('selectCurrentArray grows to a contiguous block sharing the same formula', () => {
    store.setRangeRaw(store.active, 0, 0, [['=1'], ['=1'], ['=1'], ['=2']])
    sel(1, 0)
    cmds.selectCurrentArray()
    expect(store.active.selection.ranges).toEqual([{ r1: 0, c1: 0, r2: 2, c2: 0 }])
  })

  // ---- Alt+;: visible cells only ----
  it('selectVisibleCells splits a range around hidden rows', () => {
    store.setRangeRaw(store.active, 0, 0, [['1'], ['2'], ['3'], ['4']])
    store.updateLayout(store.active, 'Hide Rows', (l) => { l.hiddenRows = { 1: true } })
    sel(0, 0, 3, 0)
    cmds.selectVisibleCells()
    const ranges = store.active.selection.ranges.map((r) => ({ ...r }))
    expect(ranges).toEqual([
      { r1: 0, c1: 0, r2: 0, c2: 0 },
      { r1: 2, c1: 0, r2: 3, c2: 0 }
    ])
  })

  // ---- Alt+Down: pick-from-list ----
  it('pickListValues returns unique values from the column above/below, skipping formulas', () => {
    store.setRangeRaw(store.active, 0, 0, [['Apple'], ['Banana'], ['Apple'], ['=1+1'], [''], ['Cherry']])
    sel(4, 0) // the blank row between the two blocks
    const values = cmds.pickListValues()
    expect(values).toEqual(['Apple', 'Banana', 'Cherry'])
  })
})

describe('clipboard helpers', () => {
  it('parses Excel clipboard text with quoted fields', async () => {
    const { parseTsv } = await import('./commands')
    expect(parseTsv('a\tb\r\n"x\ny"\t"say ""hi"""\r\n')).toEqual([['a', 'b'], ['x\ny', 'say "hi"']])
    expect(parseTsv('=A1+B1\n')).toEqual([['=A1+B1']])
  })
  it('tiles a block across a selection that is an exact multiple of it', async () => {
    const { tileTargets } = await import('./commands')
    expect(tileTargets({ r1: 0, c1: 0, r2: 3, c2: 0 }, 1, 1, false).length).toBe(4)
    expect(tileTargets({ r1: 0, c1: 0, r2: 3, c2: 1 }, 2, 1, false).length).toBe(4)
    expect(tileTargets({ r1: 0, c1: 0, r2: 2, c2: 0 }, 2, 1, false)).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 0 }])
  })
})

describe('clipboard with real Excel', () => {
  beforeEach(() => store.newWorkbook())
  const g = globalThis as { window?: unknown }

  it('reads which cells Excel marked as Text (class or inline style)', () => {
    // Shape of Excel's own markup: classes in <style>, text cells class xl67 {mso-number-format:"\@";}
    const html = `<style>.xl65 {mso-style-parent:style0; mso-number-format:Standard;} .xl67 {mso-style-parent:style0; mso-number-format:"\\@";}</style>
      <table><tr><td class=xl65>1,234.50</td><td class=xl67>00012</td><td style='mso-number-format:"\\@"'>007</td><td>plain</td></tr></table>`
    expect(cmds.excelTextCells(html)).toEqual([[false, true, true, false]])
    expect(cmds.excelTextCells('<table><tr><td colspan=2>x</td></tr></table>')).toBeNull()
  })

  it('pasting from Excel keeps Text cells like "00012" as text', async () => {
    const prev = g.window
    g.window = { api: {
      clipboardReadText: async () => '1,234.50\t00012\r\n',
      clipboardReadHTML: async () => '<style>.xl67 {mso-number-format:"\\@";}</style><table><tr><td>1,234.50</td><td class=xl67>00012</td></tr></table>'
    } }
    try {
      sel(0, 0)
      await cmds.paste()
      expect(store.getValue(store.active, 0, 0)).toBe(1234.5)
      expect(store.getValue(store.active, 0, 1)).toBe('00012')
    } finally { g.window = prev }
  })

  it('copies text as Text and keeps line breaks in one cell, in both text and HTML', async () => {
    const prev = g.window
    let written: { text: string; html?: string } = { text: '' }
    g.window = { api: { clipboardWrite: async (d: { text: string; html?: string }) => { written = d } } }
    try {
      store.setRangeRaw(store.active, 0, 0, [["'00012", 'Line1\nLine2', '5']])
      sel(0, 0, 0, 2)
      await cmds.copy()
      expect(written.text).toBe('00012\t"Line1\nLine2"\t5')
      expect(written.html).toContain(`<td style='mso-number-format:"\\@"'>00012</td>`)
      expect(written.html).toContain('Line1<br style="mso-data-placement:same-cell;"/>Line2')
      expect(written.html).toContain("<td style=''>5</td>")
    } finally { g.window = prev }
  })
})

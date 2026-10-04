import { describe, it, expect, beforeEach } from 'vitest'
import { WorkbookStore } from './store'
import { setDayFirst } from '../engine/locale'

beforeEach(() => setDayFirst(false))

describe('store', () => {
  it('evaluates formulas and undoes', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, '2')
    s.setRaw(sh, 1, 0, '3')
    s.setRaw(sh, 2, 0, '=SUM(A1:A2)')
    expect(s.getValue(sh, 2, 0)).toBe(5)
    s.undo()
    expect(s.getRaw(sh, 2, 0)).toBe('')
    s.redo()
    s.insertRows(sh, 0, 1)
    expect(s.getRaw(sh, 3, 0)).toBe('=SUM(A2:A3)')
    s.undo()
    expect(s.getRaw(sh, 2, 0)).toBe('=SUM(A1:A2)')
    s.clear(sh, [{ r1: 0, c1: 0, r2: 1, c2: 0 }])
    expect(s.getValue(sh, 2, 0)).toBe(0)
    s.undo()
    expect(s.getValue(sh, 2, 0)).toBe(5)
    s.addSheet()
    expect(s.sheets.map((x) => x.name)).toEqual(['Sheet1', 'Sheet2'])
    s.undo()
    expect(s.sheets.length).toBe(1)
  })
})

describe('input parsing', () => {
  it('parses US dates', () => {
    const s = new WorkbookStore()
    s.setRaw(s.active, 0, 0, '1/31/2025')
    expect(s.getValue(s.active, 0, 0)).toBe(45688)
  })
  it('stores a typed "1,234.50" as the number 1234.5 (it was text that SUM skipped)', () => {
    const s = new WorkbookStore()
    s.setRaw(s.active, 0, 0, '1,234.50')
    s.setRaw(s.active, 1, 0, '=SUM(A1)')
    expect(s.getValue(s.active, 0, 0)).toBe(1234.5)
    expect(s.getValue(s.active, 1, 0)).toBe(1234.5)
  })
  it('keeps input in a Text-formatted cell as text ("00012", "1,234.50"), like Excel', () => {
    const s = new WorkbookStore()
    s.setCells(s.active, [{ r: 0, c: 0, raw: '', style: { numFmt: '@' } }, { r: 1, c: 0, raw: '', style: { numFmt: '@' } }], 'fmt')
    s.setRaw(s.active, 0, 0, '00012')
    s.setRaw(s.active, 1, 0, '1,234.50')
    expect(s.getValue(s.active, 0, 0)).toBe('00012')
    expect(s.getValue(s.active, 1, 0)).toBe('1,234.50')
    expect(s.getDisplay(s.active, 0, 0).text).toBe('00012')
  })
  it('does not save a spilled result as data (it froze FILTER output and dropped leading zeros)', () => {
    const s = new WorkbookStore()
    s.setRaw(s.active, 0, 0, "'00012")
    s.setRaw(s.active, 1, 0, "'00056")
    s.setRaw(s.active, 0, 2, '=FILTER(A1:A2,A1:A2<>"")')
    expect(s.getValue(s.active, 1, 2)).toBe('00056')
    const cells = s.toData().sheets[0].cells
    expect(cells['0,2'].raw).toBe('=FILTER(A1:A2,A1:A2<>"")')
    expect(cells['1,2']).toBeUndefined()
  })
  it('parses day-first dates on an English (India) PC, like Excel', () => {
    setDayFirst(true)
    const s = new WorkbookStore()
    s.setRaw(s.active, 0, 0, '3/9/2026')
    expect(s.getValue(s.active, 0, 0)).toBe(46268)
    s.setRaw(s.active, 0, 1, '1/31/2025')
    expect(s.getValue(s.active, 0, 1)).toBe('1/31/2025')
  })
})

describe('AutoFilter', () => {
  function makeTable(): WorkbookStore {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRangeRaw(sh, 0, 0, [['Fruit', 'Qty'], ['Apple', '3'], ['Banana', '1'], ['Apple', '5']])
    s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }, sh)
    return s
  }

  it('toggles on over the current region and hides rows failing criteria', () => {
    const s = makeTable()
    const sh = s.active
    s.toggleAutoFilter(sh)
    expect(sh.layout.autoFilter?.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 1 })
    s.setColumnFilter(sh, 0, { selected: ['Apple'] })
    expect(sh.layout.hiddenRows[2]).toBe(true) // Banana row
    expect(sh.layout.hiddenRows[1]).toBeUndefined()
    expect(sh.layout.hiddenRows[3]).toBeUndefined()
  })

  it('undoes a filter change in one step', () => {
    const s = makeTable()
    const sh = s.active
    s.toggleAutoFilter(sh)
    s.setColumnFilter(sh, 0, { selected: ['Apple'] })
    s.undo()
    expect(sh.layout.hiddenRows[2]).toBeUndefined()
  })

  it('clearAllFilters unhides rows but keeps the dropdowns', () => {
    const s = makeTable()
    const sh = s.active
    s.toggleAutoFilter(sh)
    s.setColumnFilter(sh, 0, { selected: ['Apple'] })
    s.clearAllFilters(sh)
    expect(sh.layout.autoFilter).toBeDefined()
    expect(Object.keys(sh.layout.hiddenRows).length).toBe(0)
  })

  it('turning the filter off only unhides filter-caused rows, not manually hidden ones', () => {
    const s = makeTable()
    const sh = s.active
    s.updateLayout(sh, 'Hide', (l) => (l.hiddenRows[1] = true)) // manually hide the first data row
    s.toggleAutoFilter(sh)
    s.setColumnFilter(sh, 0, { selected: ['Apple'] }) // hides Banana row (2)
    s.toggleAutoFilter(sh) // off
    expect(sh.layout.autoFilter).toBeUndefined()
    expect(sh.layout.hiddenRows[1]).toBe(true) // still manually hidden
    expect(sh.layout.hiddenRows[2]).toBeUndefined() // filter-hidden row is back
  })

  it('keeps a manually hidden row hidden when it also fails the filter and the filter is cleared', () => {
    const s = makeTable()
    const sh = s.active
    s.updateLayout(sh, 'Hide', (l) => (l.hiddenRows[2] = true)) // Banana, hidden by hand
    s.toggleAutoFilter(sh)
    s.setColumnFilter(sh, 0, { selected: ['Apple'] }) // Banana also fails the filter
    s.clearAllFilters(sh)
    expect(sh.layout.hiddenRows[2]).toBe(true)
  })

  it('sorting shifts formula references and re-applies the filter to the moved rows', () => {
    const s = makeTable()
    const sh = s.active
    s.setRangeRaw(sh, 1, 2, [['=B2*10'], ['=B3*10'], ['=B4*10']])
    s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }, sh)
    s.toggleAutoFilter(sh)
    s.setColumnFilter(sh, 0, { selected: ['Apple'] }) // hides row 2 (Banana)
    s.sortFilterRange(sh, 1, 'asc') // by Qty: Banana 1, Apple 3, Apple 5
    expect([s.getRaw(sh, 1, 0), s.getRaw(sh, 2, 0), s.getRaw(sh, 3, 0)]).toEqual(['Banana', 'Apple', 'Apple'])
    expect([s.getRaw(sh, 1, 2), s.getRaw(sh, 2, 2), s.getRaw(sh, 3, 2)]).toEqual(['=B2*10', '=B3*10', '=B4*10'])
    expect([s.getValue(sh, 1, 2), s.getValue(sh, 2, 2), s.getValue(sh, 3, 2)]).toEqual([10, 30, 50])
    expect(sh.layout.hiddenRows[1]).toBe(true) // Banana moved to row 1 and stays filtered out
    expect(sh.layout.hiddenRows[2]).toBeUndefined()
    s.undo()
    expect(s.getRaw(sh, 2, 0)).toBe('Banana')
    expect(sh.layout.hiddenRows[2]).toBe(true)
  })

  it('sorts rows with a given fill colour to the top, keeping the others in order', () => {
    const s = makeTable()
    const sh = s.active
    s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }, sh)
    s.toggleAutoFilter(sh)
    s.applyStyle(sh, [{ r1: 3, c1: 0, r2: 3, c2: 0 }], { fill: '#ff0000' })
    const before = [s.getRaw(sh, 1, 0), s.getRaw(sh, 2, 0), s.getRaw(sh, 3, 0)]
    s.sortFilterRange(sh, 0, 'asc', { kind: 'fill', color: '#ff0000' })
    expect(s.getRaw(sh, 1, 0)).toBe(before[2])
    expect([s.getRaw(sh, 2, 0), s.getRaw(sh, 3, 0)]).toEqual([before[0], before[1]])
  })

  it('finds a current region larger than 10,000 rows', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRangeRaw(sh, 0, 0, Array.from({ length: 12000 }, (_, i) => [String(i)]))
    expect(s.currentRegion(sh, 0, 0)).toEqual({ r1: 0, c1: 0, r2: 11999, c2: 0 })
  })
})

describe('Notes & Comments', () => {
  it('sets, edits and deletes a note as one undo step each', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setNote(sh, 0, 0, 'Check this value', 'Ann')
    expect(sh.layout.notes['0,0']).toEqual({ text: 'Check this value', author: 'Ann' })
    s.undo()
    expect(sh.layout.notes['0,0']).toBeUndefined()
    s.redo()
    s.deleteNote(sh, 0, 0)
    expect(sh.layout.notes['0,0']).toBeUndefined()
  })

  it('adds a comment thread and a reply, then resolves it', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.addComment(sh, 1, 1, 'What does this mean?', 'Ann')
    s.replyToComment(sh, 1, 1, 'It means X', 'Bo')
    expect(sh.layout.comments['1,1'].replies).toHaveLength(1)
    s.resolveComment(sh, 1, 1, true)
    expect(sh.layout.comments['1,1'].resolved).toBe(true)
  })

  it('convertNotesToComments moves every note into a comment thread', () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setNote(sh, 0, 0, 'A note', 'Ann')
    s.convertNotesToComments(sh)
    expect(sh.layout.notes['0,0']).toBeUndefined()
    expect(sh.layout.comments['0,0'].text).toBe('A note')
  })
})

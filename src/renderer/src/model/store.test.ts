import { describe, it, expect } from 'vitest'
import { WorkbookStore } from './store'

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

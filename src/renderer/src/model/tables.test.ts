import { describe, it, expect } from 'vitest'
import { WorkbookStore } from './store'
import {
  createTable, resizeTable, convertToRange, setTableOption, setColumnTotalsFunction,
  maybeAutoExpandTable, maybeAutoFillCalculatedColumn, tabAddsRow, tableAt, displayTotalsFormula
} from './tables'

function makeSheetWithData(store: WorkbookStore) {
  const sheet = store.active
  store.setCells(sheet, [
    { r: 0, c: 0, raw: 'Name' }, { r: 0, c: 1, raw: 'Qty' }, { r: 0, c: 2, raw: 'Price' },
    { r: 1, c: 0, raw: 'Apple' }, { r: 1, c: 1, raw: '2' }, { r: 1, c: 2, raw: '1.5' },
    { r: 2, c: 0, raw: 'Pear' }, { r: 2, c: 1, raw: '3' }, { r: 2, c: 2, raw: '2' }
  ], 'Test')
  return sheet
}

describe('createTable', () => {
  it('adds Column1..N headers when the selection has none', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 0, raw: 'a' }, { r: 0, c: 1, raw: '1' },
      { r: 1, c: 0, raw: 'b' }, { r: 1, c: 1, raw: '2' }
    ], 'Test')
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }, false)
    expect(t.columns.map((c) => c.name)).toEqual(['Column1', 'Column2'])
    // A header row is inserted above the original data.
    expect(store.getRaw(sheet, 0, 0)).toBe('Column1')
    expect(store.getRaw(sheet, 1, 0)).toBe('a')
    expect(t.name).toBe('Table1')
  })

  it('reuses the first row as headers, adds an AutoFilter, and bolds the header row', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    expect(t.columns.map((c) => c.name)).toEqual(['Name', 'Qty', 'Price'])
    expect(sheet.layout.autoFilter?.range).toEqual({ r1: 0, c1: 0, r2: 2, c2: 2 })
    expect(store.getStyle(sheet, 0, 0)?.bold).toBe(true)
  })

  it('dedupes duplicate header text like Excel (Column2, Column22, ...)', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [{ r: 0, c: 0, raw: 'X' }, { r: 0, c: 1, raw: 'X' }, { r: 1, c: 0, raw: '1' }, { r: 1, c: 1, raw: '2' }], 'Test')
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }, true)
    expect(t.columns.map((c) => c.name)).toEqual(['X', 'X2'])
  })
})

describe('table auto-expansion', () => {
  it('grows the table when typing directly below the last row', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    store.setRaw(sheet, 3, 0, 'Plum')
    maybeAutoExpandTable(store, sheet, 3, 0, 'Plum')
    const grown = tableAt(sheet, 3, 0)
    expect(grown?.id).toBe(t.id)
    expect(grown?.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 })
    expect(sheet.layout.autoFilter?.range.r2).toBe(3)
  })

  it('grows the table when typing directly to the right of the last column', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    store.setRaw(sheet, 1, 3, 'Extra')
    maybeAutoExpandTable(store, sheet, 1, 3, 'Extra')
    const grown = tableAt(sheet, 1, 3)
    expect(grown?.range.c2).toBe(3)
    expect(grown?.columns.length).toBe(4)
    expect(grown?.columns[3].name).toBe('Column4')
  })

  it('does not expand when the row/column is not adjacent to the table', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    store.setRaw(sheet, 5, 0, 'far away')
    maybeAutoExpandTable(store, sheet, 5, 0, 'far away')
    expect(tableAt(sheet, 0, 0)?.range).toEqual(t.range)
  })
})

describe('Total Row', () => {
  it('writes a real SUBTOTAL formula but displays a structured reference', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    setTableOption(store, sheet, t, 'totalRow', true)
    const t2 = tableAt(sheet, 0, 0)!
    setColumnTotalsFunction(store, sheet, t2, 1, 'sum')
    // Engine sees a normal A1 SUBTOTAL formula over the data rows (excludes header & total row).
    expect(store.getRaw(sheet, 3, 1)).toBe('=SUBTOTAL(109,B2:B3)')
    expect(store.getValue(sheet, 3, 1)).toBe(5) // 2 + 3
    // Formula bar / Show Formulas show the structured-reference form instead.
    expect(displayTotalsFormula(sheet, 3, 1)).toBe('=SUBTOTAL(109,[Qty])')
  })

  it('Table Style Options > Total Row grows/shrinks the table by exactly one row', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    setTableOption(store, sheet, t, 'totalRow', true)
    expect(tableAt(sheet, 0, 0)?.range.r2).toBe(3)
    expect(store.getRaw(sheet, 3, 0)).toBe('Total')
    const t2 = tableAt(sheet, 0, 0)!
    setTableOption(store, sheet, t2, 'totalRow', false)
    expect(tableAt(sheet, 0, 0)?.range.r2).toBe(2)
  })
})

describe('Tab adds a new row', () => {
  it('Tab in the last cell of the last row grows the table by one row', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    const next = tabAddsRow(store, sheet, 2, 2)
    expect(next).toEqual({ r: 3, c: 0 })
    expect(tableAt(sheet, 3, 0)?.id).toBe(t.id)
    expect(tableAt(sheet, 3, 0)?.range.r2).toBe(3)
  })

  it('does nothing when not in the last cell', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    expect(tabAddsRow(store, sheet, 1, 2)).toBeNull()
  })

  it('with a Total Row, inserts a fresh data row above it instead of growing into it', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    setTableOption(store, sheet, t, 'totalRow', true)
    const t2 = tableAt(sheet, 0, 0)!
    const next = tabAddsRow(store, sheet, t2.range.r2 - 1, 2)
    expect(next).toEqual({ r: 3, c: 0 })
    const t3 = tableAt(sheet, 0, 0)!
    expect(t3.range.r2).toBe(4) // one more data row + the total row, still at the bottom
    expect(t3.options.totalRow).toBe(true)
  })
})

describe('calculated columns', () => {
  it('auto-fills a formula typed in one data cell down the rest of the column', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    // Add a 4th (blank) column, then type a formula into its first data row.
    store.setRaw(sheet, 0, 3, 'Total')
    maybeAutoExpandTable(store, sheet, 0, 3, 'Total')
    store.setRaw(sheet, 1, 3, '=B2*C2')
    maybeAutoFillCalculatedColumn(store, sheet, 1, 3, '=B2*C2')
    expect(store.getRaw(sheet, 2, 3)).toBe('=B3*C3')
    expect(store.getValue(sheet, 2, 3)).toBe(6) // 3 * 2
  })

  it('does not overwrite a column that already holds different data', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    // Price column (C) already has independent values in every row - a formula typed into one
    // cell must not clobber the others.
    store.setRaw(sheet, 1, 2, '=B1*2')
    maybeAutoFillCalculatedColumn(store, sheet, 1, 2, '=B1*2')
    expect(store.getRaw(sheet, 2, 2)).toBe('2') // untouched
  })
})

describe('convertToRange', () => {
  it('drops the table object and its AutoFilter but keeps every cell value', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    convertToRange(store, sheet, t)
    expect(sheet.layout.tables?.length ?? 0).toBe(0)
    expect(sheet.layout.autoFilter).toBeUndefined()
    expect(store.getRaw(sheet, 1, 0)).toBe('Apple')
    expect(store.getRaw(sheet, 0, 0)).toBe('Name')
  })
})

describe('resizeTable', () => {
  it('grows the range and adds Column4 with a default name', () => {
    const store = new WorkbookStore()
    const sheet = makeSheetWithData(store)
    const t = createTable(store, sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    resizeTable(store, sheet, t, { r1: 0, c1: 0, r2: 2, c2: 3 })
    const t2 = tableAt(sheet, 0, 0)!
    expect(t2.range.c2).toBe(3)
    expect(t2.columns.length).toBe(4)
    expect(t2.columns[3].name).toBe('Column4')
    expect(t2.columns[0].name).toBe('Name') // existing columns kept
  })
})

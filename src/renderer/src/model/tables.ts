// Excel Tables (ListObjects): create/resize/rename/convert-to-range, Total Row SUBTOTAL formulas,
// calculated-column autofill and table-region auto-expand. Kept separate from store.ts (per the
// project convention: table CRUD takes the store as a parameter, same shape as data/removeDuplicates.ts)
// so store.ts itself only needs a couple of small hooks.
import { WorkbookStore, SheetState } from './store'
import {
  Range, Table, TableColumn, TableStyleOptions, CellStyle, normRange, rangeA1, colName,
  defaultTableStyleOptions, tableDataRange, TotalsRowFunction, TOTALS_SUBTOTAL_CODE
} from './types'
import { DEFAULT_TABLE_STYLE, tableStyleByName } from './tableStyles'
import { shiftFormula } from '../engine/shiftFormula'

let nextTableSeq = 1

/** Excel's default new-table name: Table1, Table2, ... unique across the whole workbook. */
export function nextTableName(store: WorkbookStore): string {
  const used = new Set<string>()
  for (const s of store.sheets) for (const t of s.layout.tables ?? []) used.add(t.name.toLowerCase())
  let n = nextTableSeq
  while (used.has(('table' + n).toLowerCase())) n++
  nextTableSeq = n + 1
  return 'Table' + n
}

export function allTables(store: WorkbookStore): { sheet: SheetState; table: Table }[] {
  const out: { sheet: SheetState; table: Table }[] = []
  for (const s of store.sheets) for (const t of s.layout.tables ?? []) out.push({ sheet: s, table: t })
  return out
}

/** The table containing (r,c) on `sheet`, if any. */
export function tableAt(sheet: SheetState, r: number, c: number): Table | undefined {
  return (sheet.layout.tables ?? []).find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2)
}

/** Excel's name-validity rules for a table (same character/shape rules as a defined name). */
export function validateTableName(store: WorkbookStore, name: string, exceptId?: string): string | null {
  const trimmed = name.trim()
  if (!trimmed) return 'The syntax of this name isn\'t correct. Try again with valid syntax.'
  if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(trimmed)) return 'The syntax of this name isn\'t correct. Try again with valid syntax.'
  if (/^[A-Z]{1,3}\d{1,7}$/i.test(trimmed)) return 'The name that you specified looks like a cell reference. Try a different name.'
  for (const { table } of allTables(store)) {
    if (table.id !== exceptId && table.name.toLowerCase() === trimmed.toLowerCase())
      return 'The table name already exists. Enter a different name for the table.'
  }
  return null
}

function uniqueColumnNames(raws: string[]): string[] {
  const used = new Map<string, number>()
  return raws.map((raw, i) => {
    let base = raw.trim() || 'Column' + (i + 1)
    let name = base
    while ([...used.keys()].includes(name.toLowerCase())) {
      const n = (used.get(base.toLowerCase()) ?? 1) + 1
      used.set(base.toLowerCase(), n)
      name = base + n
    }
    used.set(name.toLowerCase(), used.get(name.toLowerCase()) ?? 1)
    return name
  })
}

/** Home > Format as Table / Insert > Table / Ctrl+T / Ctrl+L. Adds an AutoFilter over the header
 *  row (reusing the sheet's existing AutoFilter machinery - see the report for the one-table-per-
 *  sheet-AutoFilter tradeoff this implies) and Column1..N headers when `hasHeaders` is false. */
export function createTable(store: WorkbookStore, sheet: SheetState, range: Range, hasHeaders: boolean, styleName = DEFAULT_TABLE_STYLE): Table {
  const n = normRange(range)
  const headerNames = uniqueColumnNames(
    Array.from({ length: n.c2 - n.c1 + 1 }, (_, i) => (hasHeaders ? store.getRaw(sheet, n.r1, n.c1 + i) : ''))
  )
  const table: Table = {
    id: 'tbl' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    name: nextTableName(store),
    range: n,
    styleName,
    options: { ...defaultTableStyleOptions(), headerRow: true },
    columns: headerNames.map((name) => ({ name }))
  }
  store.transact('Insert Table', () => {
    if (!hasHeaders) {
      // Excel inserts a header row above the selection when the data has none.
      store.insertRows(sheet, n.r1, 1)
      table.range = { r1: n.r1, c1: n.c1, r2: n.r2 + 1, c2: n.c2 }
    }
    const cells = headerNames.map((name, i) => ({ r: table.range.r1, c: n.c1 + i, raw: name }))
    store.setCells(sheet, cells, 'Insert Table')
    store.applyStyle(sheet, [{ r1: table.range.r1, c1: table.range.c1, r2: table.range.r1, c2: table.range.c2 }], { bold: true }, 'Insert Table')
    store.updateLayout(sheet, 'Insert Table', (l) => {
      l.tables = [...(l.tables ?? []), table]
      // One sheet-wide AutoFilter, reused for the table's header dropdowns (see report).
      if (!l.autoFilter) l.autoFilter = { range: { r1: table.range.r1, c1: table.range.c1, r2: table.range.r2, c2: table.range.c2 }, columns: {}, filteredOutRows: [] }
    })
  })
  return table
}

function updateTable(store: WorkbookStore, sheet: SheetState, id: string, label: string, fn: (t: Table) => void): void {
  store.updateLayout(sheet, label, (l) => {
    const t = (l.tables ?? []).find((x) => x.id === id)
    if (t) fn(t)
  })
}

export function renameTable(store: WorkbookStore, sheet: SheetState, table: Table, name: string): string | null {
  const err = validateTableName(store, name, table.id)
  if (err) return err
  updateTable(store, sheet, table.id, 'Rename Table', (t) => (t.name = name.trim()))
  return null
}

export function setTableStyle(store: WorkbookStore, sheet: SheetState, table: Table, styleName: string): void {
  updateTable(store, sheet, table.id, 'Table Style', (t) => (t.styleName = styleName))
}

export function setTableOption(store: WorkbookStore, sheet: SheetState, table: Table, key: keyof TableStyleOptions, value: boolean): void {
  store.transact('Table Style Options', () => {
    if (key === 'filterButton') {
      // This app shares one AutoFilter object per sheet across every table on it (see report), so
      // toggling the Filter Button checkbox adds/removes that shared AutoFilter outright.
      store.updateLayout(sheet, 'Table Style Options', (l) => {
        const t = (l.tables ?? []).find((x) => x.id === table.id)
        if (t) t.options.filterButton = value
        const dataRange = { r1: table.range.r1, c1: table.range.c1, r2: table.options.totalRow ? table.range.r2 - 1 : table.range.r2, c2: table.range.c2 }
        if (value) l.autoFilter = { range: dataRange, columns: l.autoFilter?.columns ?? {}, filteredOutRows: l.autoFilter?.filteredOutRows ?? [] }
        else if (l.autoFilter) { for (const r of l.autoFilter.filteredOutRows) delete l.hiddenRows[r]; delete l.autoFilter }
      })
      return
    }
    if (key === 'totalRow') {
      // Grow/shrink the table range by exactly one row for the Total Row, and seed/clear its cells.
      const wasOn = table.options.totalRow
      if (value && !wasOn) {
        const totalR = table.range.r2 + 1
        store.updateLayout(sheet, 'Total Row', (l) => {
          const t = (l.tables ?? []).find((x) => x.id === table.id)
          if (!t) return
          t.range = { ...t.range, r2: totalR }
          t.options.totalRow = true
        })
        const cells = table.columns.map((_, i) => ({ r: totalR, c: table.range.c1 + i, raw: i === 0 ? 'Total' : '' }))
        store.setCells(sheet, cells, 'Total Row')
        store.applyStyle(sheet, [{ r1: totalR, c1: table.range.c1, r2: totalR, c2: table.range.c2 }], { bold: true }, 'Total Row')
      } else if (!value && wasOn) {
        const totalR = table.range.r2
        const cells = table.columns.map((_, i) => ({ r: totalR, c: table.range.c1 + i, raw: '' }))
        store.setCells(sheet, cells, 'Total Row')
        store.updateLayout(sheet, 'Total Row', (l) => {
          const t = (l.tables ?? []).find((x) => x.id === table.id)
          if (!t) return
          t.range = { ...t.range, r2: totalR - 1 }
          t.options.totalRow = false
        })
      }
      return
    }
    updateTable(store, sheet, table.id, 'Table Style Options', (t) => { t.options[key] = value })
  })
}

/** Total Row dropdown: sets column `colIdx`'s aggregate and writes the real A1 SUBTOTAL formula
 *  into the engine (HyperFormula has no structured references - see displayTotalsFormula below for
 *  the `[Column]` text shown in the formula bar instead). */
export function setColumnTotalsFunction(store: WorkbookStore, sheet: SheetState, table: Table, colIdx: number, fn: TotalsRowFunction, customFormula?: string): void {
  if (!table.options.totalRow) return
  const totalR = table.range.r2
  const c = table.range.c1 + colIdx
  const data = tableDataRange(table)
  store.transact('Total Row Function', () => {
    updateTable(store, sheet, table.id, 'Total Row Function', (t) => {
      const col = t.columns[colIdx]
      col.totalsRowFunction = fn
      col.totalsRowFormula = fn === 'custom' ? customFormula : undefined
      col.totalsRowLabel = undefined
    })
    let raw = ''
    if (fn === 'custom' && customFormula) raw = customFormula
    else if (fn !== 'none') {
      const code = TOTALS_SUBTOTAL_CODE[fn as Exclude<TotalsRowFunction, 'none' | 'custom'>]
      raw = `=SUBTOTAL(${code},${rangeA1({ r1: data.r1, c1: c, r2: data.r2, c2: c })})`
    }
    store.setRaw(sheet, totalR, c, raw)
  })
}

export function setTotalsRowLabel(store: WorkbookStore, sheet: SheetState, table: Table, colIdx: number, label: string): void {
  const totalR = table.range.r2
  const c = table.range.c1 + colIdx
  store.transact('Total Row', () => {
    updateTable(store, sheet, table.id, 'Total Row', (t) => {
      t.columns[colIdx].totalsRowFunction = 'none'
      t.columns[colIdx].totalsRowLabel = label
    })
    store.setRaw(sheet, totalR, c, label)
  })
}

/** Formula-bar/Show-Formulas display text for a Total Row cell: `=SUBTOTAL(109,[Column])` instead
 *  of the real range the engine holds. Returns null for any other cell (caller falls back to the
 *  real raw text). */
export function displayTotalsFormula(sheet: SheetState, r: number, c: number): string | null {
  const t = tableAt(sheet, r, c)
  if (!t || t.range.r2 !== r || !t.options.totalRow) return null
  const col = t.columns[c - t.range.c1]
  if (!col || !col.totalsRowFunction || col.totalsRowFunction === 'none') return null
  if (col.totalsRowFunction === 'custom') return col.totalsRowFormula ?? null
  const code = TOTALS_SUBTOTAL_CODE[col.totalsRowFunction as Exclude<TotalsRowFunction, 'none' | 'custom'>]
  return `=SUBTOTAL(${code},[${col.name}])`
}

/** Data > Convert to Range: drops the table object (and its filter/banding) but leaves every cell
 *  value and format exactly as it is. If the table's AutoFilter is the sheet's shared one, it's
 *  removed too (unhiding any rows it had filtered out), matching Excel dropping the filter buttons. */
export function convertToRange(store: WorkbookStore, sheet: SheetState, table: Table): void {
  store.updateLayout(sheet, 'Convert to Range', (l) => {
    l.tables = (l.tables ?? []).filter((t) => t.id !== table.id)
    if (l.autoFilter && rangesEqualish(l.autoFilter.range, { r1: table.range.r1, c1: table.range.c1, r2: table.options.totalRow ? table.range.r2 - 1 : table.range.r2, c2: table.range.c2 })) {
      for (const r of l.autoFilter.filteredOutRows) delete l.hiddenRows[r]
      delete l.autoFilter
    }
  })
}

function rangesEqualish(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2
}

export function resizeTable(store: WorkbookStore, sheet: SheetState, table: Table, newRange: Range): void {
  const n = normRange(newRange)
  store.transact('Resize Table', () => {
    updateTable(store, sheet, table.id, 'Resize Table', (t) => {
      const oldC1 = t.range.c1, oldC2 = t.range.c2
      const width = n.c2 - n.c1 + 1
      const cols: TableColumn[] = []
      for (let i = 0; i < width; i++) {
        const c = n.c1 + i
        cols.push(c >= oldC1 && c <= oldC2 ? t.columns[c - oldC1] : { name: 'Column' + (i + 1) })
      }
      t.range = n
      t.columns = cols
    })
    // Keep the shared AutoFilter's range in step with the header/data extent.
    store.updateLayout(sheet, 'Resize Table', (l) => {
      const t = (l.tables ?? []).find((x) => x.id === table.id)
      if (!t || !l.autoFilter) return
      l.autoFilter.range = { r1: t.range.r1, c1: t.range.c1, r2: t.options.totalRow ? t.range.r2 - 1 : t.range.r2, c2: t.range.c2 }
    })
  })
}

// ---------- typing hooks (called from chrome/ui.ts commitEdit, after the raw write lands) ----------

/** Excel auto-expands a table when you type directly below its last row or to the right of its
 *  last column (within the table's row span), with a small "Undo Table AutoExpansion" smart tag.
 *  Here the expansion is folded into the SAME undo step as the keystroke that triggered it, so one
 *  Ctrl+Z removes both - the closest equivalent this app's undo model has to that smart tag. */
export function maybeAutoExpandTable(store: WorkbookStore, sheet: SheetState, r: number, c: number, raw: string): void {
  if (!raw) return
  for (const t of sheet.layout.tables ?? []) {
    if (r === t.range.r2 + 1 && c >= t.range.c1 && c <= t.range.c2 && !t.options.totalRow) {
      updateTable(store, sheet, t.id, 'Table AutoExpansion', (tt) => (tt.range = { ...tt.range, r2: r }))
      store.updateLayout(sheet, 'Table AutoExpansion', (l) => {
        if (l.autoFilter && rangesEqualish(l.autoFilter.range, { ...t.range, r2: t.range.r2 })) l.autoFilter.range = { ...l.autoFilter.range, r2: r }
      })
      return
    }
    if (c === t.range.c2 + 1 && r >= t.range.r1 && r <= (t.options.totalRow ? t.range.r2 - 1 : t.range.r2)) {
      updateTable(store, sheet, t.id, 'Table AutoExpansion', (tt) => { tt.range = { ...tt.range, c2: c }; tt.columns = [...tt.columns, { name: 'Column' + (tt.columns.length + 1) }] })
      store.updateLayout(sheet, 'Table AutoExpansion', (l) => {
        if (l.autoFilter) l.autoFilter.range = { ...l.autoFilter.range, c2: Math.max(l.autoFilter.range.c2, c) }
      })
      return
    }
  }
}

/** Calculated columns: a formula typed anywhere in a table's data column fills the rest of that
 *  column (shifted per row, like fill-down), as long as the rest of the column is blank or already
 *  holds the same calculated formula - so it won't clobber genuinely different per-row data. */
export function maybeAutoFillCalculatedColumn(store: WorkbookStore, sheet: SheetState, r: number, c: number, raw: string): void {
  const t = tableAt(sheet, r, c)
  if (!t || !raw.startsWith('=')) return
  const data = tableDataRange(t)
  if (r < data.r1 || r > data.r2) return
  const colIdx = c - t.range.c1
  const col = t.columns[colIdx]
  const shifted = (targetR: number) => shiftFormula(raw, targetR - r, 0)
  const canOverwrite = (rr: number): boolean => {
    if (rr === r) return true
    const existing = store.getRaw(sheet, rr, c)
    return existing === '' || existing === shiftFormula(col.calculatedFormula ?? '', rr - r, 0)
  }
  let allOk = true
  for (let rr = data.r1; rr <= data.r2; rr++) if (!canOverwrite(rr)) { allOk = false; break }
  if (!allOk) return
  const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
  for (let rr = data.r1; rr <= data.r2; rr++) if (rr !== r) cells.push({ r: rr, c, raw: shifted(rr), style: store.getStyle(sheet, rr, c) })
  if (cells.length) store.setCells(sheet, cells, 'Calculated Column')
  updateTable(store, sheet, t.id, 'Calculated Column', (tt) => (tt.columns[colIdx].calculatedFormula = raw))
}

/** Tab pressed in the last cell of a table's last data row: adds a new row (like Excel), and
 *  returns the cell the active selection should move to, or null if this isn't that case. */
export function tabAddsRow(store: WorkbookStore, sheet: SheetState, r: number, c: number): { r: number; c: number } | null {
  const t = tableAt(sheet, r, c)
  if (!t) return null
  const data = tableDataRange(t)
  if (r !== data.r2 || c !== t.range.c2) return null
  const newR = data.r2 + 1
  if (t.options.totalRow) {
    // Insert a fresh data row above the Total Row instead of growing into it.
    store.transact('Table AutoExpansion', () => {
      store.insertRows(sheet, newR, 1)
      updateTable(store, sheet, t.id, 'Table AutoExpansion', (tt) => (tt.range = { ...tt.range, r2: tt.range.r2 + 1 }))
    })
  } else {
    updateTable(store, sheet, t.id, 'Table AutoExpansion', (tt) => (tt.range = { ...tt.range, r2: newR }))
  }
  return { r: newR, c: t.range.c1 }
}

// ---------- rendering (Grid.tsx) ----------
export interface TableCellPaint {
  fill?: string
  color?: string
  bold?: boolean
  borderColor?: string
}

/** The banded-fill / header / total-row / first-last-column look for one table cell, computed
 *  fresh every draw from the table style + options - never written into sheet.styles, so ordinary
 *  direct cell formatting still overrides it exactly like Excel ("this fill wins unless the user
 *  applied their own"). Returns undefined for a cell with no table styling to apply. */
export function tableCellPaint(t: Table, r: number, c: number): TableCellPaint | undefined {
  const s = tableStyleByName(t.styleName)
  if (!s) return undefined
  const data = tableDataRange(t)
  if (t.options.headerRow && r === t.range.r1) {
    return { fill: s.headerFill, color: s.headerFont, bold: s.headerBold, borderColor: s.border }
  }
  if (t.options.totalRow && r === t.range.r2) {
    return { color: s.totalFont, bold: s.totalBold, borderColor: s.border }
  }
  if (r < data.r1 || r > data.r2) return undefined
  let fill: string | undefined
  let color: string | undefined
  let bold = false
  if (t.options.bandedColumns) {
    const odd = (c - t.range.c1) % 2 === 0
    fill = odd ? s.band1 : s.band2
  } else if (t.options.bandedRows) {
    const odd = (r - data.r1) % 2 === 0
    fill = odd ? s.band1 : s.band2
  }
  if (t.options.firstColumn && c === t.range.c1) { fill = s.firstColFill ?? fill; color = s.firstColFont; bold = true }
  if (t.options.lastColumn && c === t.range.c2) { fill = s.lastColFill ?? fill; color = s.lastColFont; bold = true }
  return fill || color || bold ? { fill, color, bold: bold || undefined, borderColor: s.border } : undefined
}

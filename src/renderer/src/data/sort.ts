import { WorkbookStore, SheetState } from '../model/store'
import { CellStyle, Range, isError, normRange, MAX_ROWS, MAX_COLS } from '../model/types'
import { shiftFormula } from '../engine/shiftFormula'

export interface SortLevel {
  columnIndex: number // 0-based column index
  sortOn: 'values' | 'cellColor' | 'fontColor'
  order: 'asc' | 'desc'
}

export interface SortOptions {
  hasHeaders?: boolean
  caseSensitive?: boolean
  orientation?: 'topToBottom' | 'leftToRight'
  levels: SortLevel[]
}

/** Determines type rank matching Excel: numbers(0) < text(1) < logicals(2) < errors(3) < blanks(4) */
function getTypeRank(v: any): number {
  if (v === null || v === undefined || v === '') return 4
  if (isError(v)) return 3
  if (typeof v === 'boolean') return 2
  if (typeof v === 'number') return 0
  if (typeof v === 'string') {
    // If it's an error string like #DIV/0!
    if (v.startsWith('#') && v.endsWith('!')) return 3
    return 1
  }
  return 1
}

/** Compares two values matching Excel's exact sorting rules. Blanks are always last. */
export function compareValues(a: any, b: any, order: 'asc' | 'desc', caseSensitive = false): number {
  const isBlankA = a === null || a === undefined || a === ''
  const isBlankB = b === null || b === undefined || b === ''

  if (isBlankA && isBlankB) return 0
  if (isBlankA) return 1 // blanks always at the end
  if (isBlankB) return -1

  const typeA = getTypeRank(a)
  const typeB = getTypeRank(b)

  let cmp = 0
  if (typeA !== typeB) {
    cmp = typeA - typeB
  } else {
    if (typeA === 0) {
      cmp = Number(a) - Number(b)
    } else if (typeA === 1) {
      const sa = String(a)
      const sb = String(b)
      if (caseSensitive) {
        cmp = sa < sb ? -1 : sa > sb ? 1 : 0
      } else {
        const lsa = sa.toLowerCase()
        const lsb = sb.toLowerCase()
        cmp = lsa < lsb ? -1 : lsa > lsb ? 1 : 0
        if (cmp === 0 && sa !== sb) {
          cmp = sa < sb ? -1 : 1
        }
      }
    } else if (typeA === 2) {
      const ba = Boolean(a)
      const bb = Boolean(b)
      cmp = ba === bb ? 0 : !ba ? -1 : 1
    } else if (typeA === 3) {
      const ea = isError(a) ? a.error : String(a)
      const eb = isError(b) ? b.error : String(b)
      cmp = ea.localeCompare(eb)
    }
  }

  return order === 'asc' ? cmp : -cmp
}

/** Automatically detects whether the first row of range is a header row. */
export function detectHeaders(store: WorkbookStore, sheet: SheetState, range: Range): boolean {
  const n = normRange(range)
  if (n.r2 <= n.r1) return false

  let headerStrings = 0
  let dataNonStrings = 0
  let headerFormatted = 0
  let totalCols = 0

  for (let c = n.c1; c <= n.c2; c++) {
    const val1 = store.getValue(sheet, n.r1, c)
    const val2 = store.getValue(sheet, n.r1 + 1, c)
    const raw1 = store.getRaw(sheet, n.r1, c)
    const style1 = store.getStyle(sheet, n.r1, c)
    const style2 = store.getStyle(sheet, n.r1 + 1, c)

    if (raw1 !== '') {
      totalCols++
      if (typeof val1 === 'string' && isNaN(Number(val1))) {
        headerStrings++
      }
      if (typeof val2 === 'number' || val2 instanceof Date) {
        dataNonStrings++
      }
      if (style1?.bold && !style2?.bold) {
        headerFormatted++
      }
      if (style1?.fill && !style2?.fill) {
        headerFormatted++
      }
    }
  }

  if (totalCols === 0) return false
  if (headerFormatted > 0) return true
  if (headerStrings === totalCols && dataNonStrings > 0) return true
  return false
}

/** Expands range to the contiguous populated region (CurrentRegion in Excel). */
export function findCurrentRegion(store: WorkbookStore, sheet: SheetState, range: Range): Range {
  const n = normRange(range)
  const used = store.usedRange(sheet)
  if (!used) return n

  let r1 = n.r1
  let r2 = n.r2
  let c1 = n.c1
  let c2 = n.c2

  let changed = true
  while (changed) {
    changed = false

    // Check top
    if (r1 > 0) {
      let hasData = false
      for (let c = Math.max(0, c1 - 1); c <= Math.min(MAX_COLS - 1, c2 + 1); c++) {
        if (store.getRaw(sheet, r1 - 1, c) !== '') {
          hasData = true
          break
        }
      }
      if (hasData) {
        r1--
        changed = true
      }
    }

    // Check bottom
    if (r2 < used.r2 + 1) {
      let hasData = false
      for (let c = Math.max(0, c1 - 1); c <= Math.min(MAX_COLS - 1, c2 + 1); c++) {
        if (store.getRaw(sheet, r2 + 1, c) !== '') {
          hasData = true
          break
        }
      }
      if (hasData) {
        r2++
        changed = true
      }
    }

    // Check left
    if (c1 > 0) {
      let hasData = false
      for (let r = Math.max(0, r1 - 1); r <= Math.min(MAX_ROWS - 1, r2 + 1); r++) {
        if (store.getRaw(sheet, r, c1 - 1) !== '') {
          hasData = true
          break
        }
      }
      if (hasData) {
        c1--
        changed = true
      }
    }

    // Check right
    if (c2 < used.c2 + 1) {
      let hasData = false
      for (let r = Math.max(0, r1 - 1); r <= Math.min(MAX_ROWS - 1, r2 + 1); r++) {
        if (store.getRaw(sheet, r, c2 + 1) !== '') {
          hasData = true
          break
        }
      }
      if (hasData) {
        c2++
        changed = true
      }
    }
  }

  return { r1, c1, r2, c2 }
}

/** Checks whether the selected range has adjacent populated cells outside it. */
export function hasAdjacentData(store: WorkbookStore, sheet: SheetState, range: Range): boolean {
  const n = normRange(range)
  const expanded = findCurrentRegion(store, sheet, n)
  return (
    expanded.r1 < n.r1 ||
    expanded.r2 > n.r2 ||
    expanded.c1 < n.c1 ||
    expanded.c2 > n.c2
  )
}

interface RowItem {
  origRow: number
  cells: { col: number; raw: string; value: any; style?: CellStyle }[]
}

/** Sorts rows in range with multi-level criteria, preserving styles and updating formulas. */
export function sortRange(
  store: WorkbookStore,
  sheet: SheetState,
  range: Range,
  options: SortOptions
): void {
  const n = normRange(range)
  const hasHeaders = options.hasHeaders ?? detectHeaders(store, sheet, n)
  const startRow = hasHeaders ? n.r1 + 1 : n.r1

  if (startRow >= n.r2) return // Nothing or only 1 row to sort

  const rows: RowItem[] = []
  for (let r = startRow; r <= n.r2; r++) {
    const cells: RowItem['cells'] = []
    for (let c = n.c1; c <= n.c2; c++) {
      cells.push({
        col: c,
        raw: store.getRaw(sheet, r, c),
        value: store.getValue(sheet, r, c),
        style: store.getStyle(sheet, r, c)
      })
    }
    rows.push({ origRow: r, cells })
  }

  const levels = options.levels.length > 0 ? options.levels : [{ columnIndex: n.c1, sortOn: 'values', order: 'asc' as const }]

  rows.sort((rowA, rowB) => {
    for (const lvl of levels) {
      if (lvl.sortOn === 'cellColor') {
        const styleA = rowA.cells.find((c) => c.col === lvl.columnIndex)?.style?.fill || ''
        const styleB = rowB.cells.find((c) => c.col === lvl.columnIndex)?.style?.fill || ''
        const cmp = styleA.localeCompare(styleB)
        if (cmp !== 0) return lvl.order === 'asc' ? cmp : -cmp
      } else if (lvl.sortOn === 'fontColor') {
        const styleA = rowA.cells.find((c) => c.col === lvl.columnIndex)?.style?.color || ''
        const styleB = rowB.cells.find((c) => c.col === lvl.columnIndex)?.style?.color || ''
        const cmp = styleA.localeCompare(styleB)
        if (cmp !== 0) return lvl.order === 'asc' ? cmp : -cmp
      } else {
        const valA = rowA.cells.find((c) => c.col === lvl.columnIndex)?.value
        const valB = rowB.cells.find((c) => c.col === lvl.columnIndex)?.value
        const cmp = compareValues(valA, valB, lvl.order, options.caseSensitive)
        if (cmp !== 0) return cmp
      }
    }
    return rowA.origRow - rowB.origRow // stable sort fallback
  })

  // Apply new sorted order in a single transaction
  store.transact('Sort', () => {
    const updates: { r: number; c: number; raw: string; style?: CellStyle }[] = []
    rows.forEach((row, i) => {
      const targetRow = startRow + i
      const dr = targetRow - row.origRow
      row.cells.forEach((cell) => {
        updates.push({
          r: targetRow,
          c: cell.col,
          raw: cell.raw.startsWith('=') ? shiftFormula(cell.raw, dr, 0) : cell.raw,
          style: cell.style
        })
      })
    })

    store.setCells(sheet, updates, 'Sort')
  })
}

import { WorkbookStore, SheetState } from '../model/store'
import { Range, normRange, a1, isError } from '../model/types'

export interface FindOptions {
  query: string
  within?: 'sheet' | 'workbook' // default 'sheet'
  searchBy?: 'rows' | 'columns' // default 'rows'
  lookIn?: 'formulas' | 'values' | 'notes' | 'comments' // default 'formulas'
  matchCase?: boolean // default false
  matchEntireCell?: boolean // default false
}

export interface FindResult {
  book: string
  sheetName: string
  sheetIndex: number
  cellRef: string // e.g. '$A$1'
  r: number
  c: number
  value: string
  formula: string
}

/** Converts an Excel wildcard pattern (*, ?, ~*, ~?, ~~) into a regular expression. */
export function compileWildcardPattern(pattern: string, matchCase: boolean, matchEntire: boolean): RegExp {
  let regexStr = ''
  let i = 0
  while (i < pattern.length) {
    const ch = pattern[i]
    if (ch === '~') {
      if (i + 1 < pattern.length) {
        const next = pattern[i + 1]
        // Escape regex special chars
        regexStr += escapeRegex(next)
        i += 2
        continue
      } else {
        regexStr += escapeRegex('~')
        i++
        continue
      }
    } else if (ch === '*') {
      regexStr += '.*'
    } else if (ch === '?') {
      regexStr += '.'
    } else {
      regexStr += escapeRegex(ch)
    }
    i++
  }

  if (matchEntire) {
    regexStr = `^${regexStr}$`
  }
  return new RegExp(regexStr, matchCase ? '' : 'i')
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function cellMatches(
  text: string,
  query: string,
  matchCase: boolean,
  matchEntire: boolean
): boolean {
  if (!query) return false
  const re = compileWildcardPattern(query, matchCase, matchEntire)
  return re.test(text)
}

/** Scans a sheet in row or column order and returns all matching cells. */
export function findInSheet(
  store: WorkbookStore,
  sheet: SheetState,
  sheetIndex: number,
  options: FindOptions
): FindResult[] {
  const results: FindResult[] = []
  if (!options.query) return results

  const used = store.usedRange(sheet)
  if (!used) return results

  const matchCase = !!options.matchCase
  const matchEntire = !!options.matchEntireCell
  const lookIn = options.lookIn || 'formulas'
  const searchBy = options.searchBy || 'rows'
  const pattern = compileWildcardPattern(options.query, matchCase, matchEntire)

  const checkCell = (r: number, c: number) => {
    const raw = store.getRaw(sheet, r, c)
    const val = store.getValue(sheet, r, c)
    const display = store.getDisplay(sheet, r, c).text

    // If cell is completely empty, skip
    if (raw === '' && (val === null || val === undefined || val === '')) return

    let targetText = ''
    if (lookIn === 'formulas') {
      targetText = raw
    } else if (lookIn === 'values') {
      targetText = display
    } else {
      targetText = ''
    }

    if (pattern.test(targetText)) {
      results.push({
        book: store.fileName || 'Book1',
        sheetName: sheet.name,
        sheetIndex,
        cellRef: `$${a1(r, c).slice(0, -1)}$${r + 1}`, // e.g. $A$1
        r,
        c,
        value: display,
        formula: raw
      })
    }
  }

  if (searchBy === 'columns') {
    for (let c = used.c1; c <= used.c2; c++) {
      for (let r = used.r1; r <= used.r2; r++) {
        checkCell(r, c)
      }
    }
  } else {
    for (let r = used.r1; r <= used.r2; r++) {
      for (let c = used.c1; c <= used.c2; c++) {
        checkCell(r, c)
      }
    }
  }

  return results
}

/** Finds all matching cells in the active sheet or entire workbook. */
export function findAll(store: WorkbookStore, options: FindOptions): FindResult[] {
  if (options.within === 'workbook') {
    const all: FindResult[] = []
    store.sheets.forEach((sheet, index) => {
      all.push(...findInSheet(store, sheet, index, options))
    })
    return all
  }
  return findInSheet(store, store.active, store.activeIndex, options)
}

/** Finds the next match relative to current cell, wrapping around. */
export function findNext(
  store: WorkbookStore,
  options: FindOptions,
  current?: { sheetIndex: number; r: number; c: number }
): FindResult | null {
  const matches = findAll(store, options)
  if (!matches.length) return null

  if (!current) return matches[0]

  // Find index of first match after current position
  const searchBy = options.searchBy || 'rows'
  const isAfter = (m: FindResult): boolean => {
    if (m.sheetIndex !== current.sheetIndex) {
      return m.sheetIndex > current.sheetIndex
    }
    if (searchBy === 'columns') {
      return m.c > current.c || (m.c === current.c && m.r > current.r)
    }
    return m.r > current.r || (m.r === current.r && m.c > current.c)
  }

  for (const m of matches) {
    if (isAfter(m)) return m
  }
  return matches[0] // wrap around
}

/** Performs a Replace All across the sheet/workbook as a single undo transaction. */
export function replaceAll(
  store: WorkbookStore,
  options: FindOptions,
  replaceWith: string
): number {
  // In Excel, Replace is always performed on Formulas (raw cell contents)
  const effectiveOptions: FindOptions = {
    ...options,
    lookIn: 'formulas'
  }
  const matches = findAll(store, effectiveOptions)
  if (!matches.length) return 0

  const matchCase = !!options.matchCase
  const matchEntire = !!options.matchEntireCell
  const pattern = compileWildcardPattern(options.query, matchCase, matchEntire)

  let replacementCount = 0

  store.transact('Replace All', () => {
    // Group matches by sheet
    const bySheet = new Map<number, FindResult[]>()
    matches.forEach((m) => {
      const list = bySheet.get(m.sheetIndex) || []
      list.push(m)
      bySheet.set(m.sheetIndex, list)
    })

    bySheet.forEach((list, sheetIndex) => {
      const sheet = store.sheets[sheetIndex]
      if (!sheet) return
      const updates: { r: number; c: number; raw: string; style?: any }[] = []

      list.forEach((m) => {
        const raw = store.getRaw(sheet, m.r, m.c)
        let next = ''
        if (matchEntire) {
          next = replaceWith
        } else {
          // Replace using global regex if pattern matches multiple times
          const globalRe = new RegExp(pattern.source, pattern.flags + (pattern.flags.includes('g') ? '' : 'g'))
          next = raw.replace(globalRe, replaceWith)
        }

        if (next !== raw) {
          updates.push({
            r: m.r,
            c: m.c,
            raw: next,
            style: store.getStyle(sheet, m.r, m.c)
          })
          replacementCount++
        }
      })

      if (updates.length > 0) {
        store.setCells(sheet, updates, 'Replace All')
      }
    })
  })

  return replacementCount
}

export type SpecialType =
  | 'comments'
  | 'constants'
  | 'formulas'
  | 'blanks'
  | 'currentRegion'
  | 'currentArray'
  | 'objects'
  | 'rowDifferences'
  | 'colDifferences'
  | 'precedents'
  | 'dependents'
  | 'lastCell'
  | 'visibleCells'
  | 'conditionalFormats'
  | 'dataValidation'

export interface SpecialSubtypes {
  numbers: boolean
  text: boolean
  logicals: boolean
  errors: boolean
}

/** Finds matching cell ranges for Excel's Go To Special dialog options. */
export function getSpecialRanges(
  store: WorkbookStore,
  type: SpecialType,
  subtypes: SpecialSubtypes = { numbers: true, text: true, logicals: true, errors: true }
): Range[] {
  const sheet = store.active
  const used = store.usedRange(sheet)
  if (!used) return []

  const active = sheet.selection.active
  const selRanges = sheet.selection.ranges
  const isMultiCellSelection = selRanges.length > 1 || (
    selRanges[0] && (selRanges[0].r1 !== selRanges[0].r2 || selRanges[0].c1 !== selRanges[0].c2)
  )

  // Scope: if user selected a multi-cell range, operate within selection; otherwise entire used range
  const scopeRanges = isMultiCellSelection ? selRanges : [used]

  const matches: { r: number; c: number }[] = []

  if (type === 'lastCell') {
    return [{ r1: used.r2, c1: used.c2, r2: used.r2, c2: used.c2 }]
  }

  if (type === 'currentRegion') {
    // Expand from active cell as long as connected to non-empty cells (Excel Ctrl+A behavior)
    let r1 = active.r, r2 = active.r, c1 = active.c, c2 = active.c
    let changed = true
    while (changed) {
      changed = false
      // Check top expansion
      if (r1 > 0) {
        let hasData = false
        for (let c = Math.max(0, c1 - 1); c <= c2 + 1; c++) {
          if (store.getRaw(sheet, r1 - 1, c) !== '') { hasData = true; break }
        }
        if (hasData) { r1--; changed = true }
      }
      // Check bottom expansion
      if (r2 < used.r2 + 1) {
        let hasData = false
        for (let c = Math.max(0, c1 - 1); c <= c2 + 1; c++) {
          if (store.getRaw(sheet, r2 + 1, c) !== '') { hasData = true; break }
        }
        if (hasData) { r2++; changed = true }
      }
      // Check left expansion
      if (c1 > 0) {
        let hasData = false
        for (let r = Math.max(0, r1 - 1); r <= r2 + 1; r++) {
          if (store.getRaw(sheet, r, c1 - 1) !== '') { hasData = true; break }
        }
        if (hasData) { c1--; changed = true }
      }
      // Check right expansion
      if (c2 < used.c2 + 1) {
        let hasData = false
        for (let r = Math.max(0, r1 - 1); r <= r2 + 1; r++) {
          if (store.getRaw(sheet, r, c2 + 1) !== '') { hasData = true; break }
        }
        if (hasData) { c2++; changed = true }
      }
    }
    return [{ r1, c1, r2, c2 }]
  }

  if (type === 'visibleCells') {
    for (const rg of selRanges) {
      const n = normRange(rg)
      for (let r = n.r1; r <= n.r2; r++) {
        if (sheet.layout.hiddenRows[r]) continue
        for (let c = n.c1; c <= n.c2; c++) {
          if (sheet.layout.hiddenCols[c]) continue
          matches.push({ r, c })
        }
      }
    }
    return coordsToRanges(matches)
  }

  if (type === 'rowDifferences' || type === 'colDifferences') {
    for (const rg of selRanges) {
      const n = normRange(rg)
      if (type === 'rowDifferences') {
        const compareCol = Math.max(n.c1, Math.min(n.c2, active.c))
        for (let r = n.r1; r <= n.r2; r++) {
          const compRaw = store.getRaw(sheet, r, compareCol)
          for (let c = n.c1; c <= n.c2; c++) {
            if (c !== compareCol && store.getRaw(sheet, r, c) !== compRaw) {
              matches.push({ r, c })
            }
          }
        }
      } else {
        const compareRow = Math.max(n.r1, Math.min(n.r2, active.r))
        for (let c = n.c1; c <= n.c2; c++) {
          const compRaw = store.getRaw(sheet, compareRow, c)
          for (let r = n.r1; r <= n.r2; r++) {
            if (r !== compareRow && store.getRaw(sheet, r, c) !== compRaw) {
              matches.push({ r, c })
            }
          }
        }
      }
    }
    return coordsToRanges(matches)
  }

  for (const rg of scopeRanges) {
    const n = normRange(rg)
    for (let r = n.r1; r <= n.r2; r++) {
      for (let c = n.c1; c <= n.c2; c++) {
        const raw = store.getRaw(sheet, r, c)
        const val = store.getValue(sheet, r, c)

        if (type === 'blanks') {
          if (raw === '' && (val === null || val === undefined || val === '')) {
            matches.push({ r, c })
          }
        } else if (type === 'constants') {
          if (raw !== '' && !raw.startsWith('=')) {
            if (isSubtypeMatch(val, subtypes)) {
              matches.push({ r, c })
            }
          }
        } else if (type === 'formulas') {
          if (raw.startsWith('=')) {
            if (isSubtypeMatch(val, subtypes)) {
              matches.push({ r, c })
            }
          }
        }
      }
    }
  }

  return coordsToRanges(matches)
}

function isSubtypeMatch(val: unknown, subtypes: SpecialSubtypes): boolean {
  if (isError(val as any)) return subtypes.errors
  if (typeof val === 'number') return subtypes.numbers
  if (typeof val === 'string') return subtypes.text
  if (typeof val === 'boolean') return subtypes.logicals
  return false
}

/** Consolidates a list of cell coordinates into Range objects. */
export function coordsToRanges(coords: { r: number; c: number }[]): Range[] {
  if (!coords.length) return []
  // Sort by row, then col
  const sorted = [...coords].sort((a, b) => a.r - b.r || a.c - b.c)
  const ranges: Range[] = []
  sorted.forEach((pt) => {
    ranges.push({ r1: pt.r, c1: pt.c, r2: pt.r, c2: pt.c })
  })
  return ranges
}

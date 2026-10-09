import { WorkbookStore, SheetState } from '../model/store'
import { Range, normRange, CellStyle } from '../model/types'
import { shiftFormula } from '../engine/shiftFormula'

export interface RemoveDuplicatesResult {
  removedCount: number
  remainingCount: number
}

export function removeDuplicates(
  store: WorkbookStore,
  sheet: SheetState,
  range: Range,
  columnIndices: number[],
  hasHeaders: boolean
): RemoveDuplicatesResult {
  const n = normRange(range)
  const startRow = hasHeaders ? n.r1 + 1 : n.r1
  if (startRow > n.r2) {
    return { removedCount: 0, remainingCount: 0 }
  }

  const seenKeys = new Set<string>()
  const uniqueRows: { origRow: number; cells: { col: number; raw: string; style?: CellStyle }[] }[] = []
  let removedCount = 0

  const read = store.rawReader(sheet, (n.r2 - startRow + 1) * (n.c2 - n.c1 + 1))
  for (let r = startRow; r <= n.r2; r++) {
    const keyParts = columnIndices.map((col) => {
      const v = store.getValue(sheet, r, col)
      return String(v ?? read(r, col)).toLowerCase()
    })
    const compositeKey = keyParts.join('\x1F')

    if (seenKeys.has(compositeKey)) {
      removedCount++
    } else {
      seenKeys.add(compositeKey)
      const cells: { col: number; raw: string; style?: CellStyle }[] = []
      for (let c = n.c1; c <= n.c2; c++) {
        cells.push({
          col: c,
          raw: read(r, c),
          style: store.getStyle(sheet, r, c)
        })
      }
      uniqueRows.push({ origRow: r, cells })
    }
  }

  if (removedCount > 0) {
    store.transact('Remove Duplicates', () => {
      const updates: { r: number; c: number; raw: string; style?: CellStyle }[] = []

      // 1. Write unique rows to the top of data section
      uniqueRows.forEach((row, i) => {
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

      // 2. Clear remaining rows at the bottom of the range
      const endUniqueRow = startRow + uniqueRows.length
      for (let r = endUniqueRow; r <= n.r2; r++) {
        for (let c = n.c1; c <= n.c2; c++) {
          updates.push({
            r,
            c,
            raw: '',
            style: undefined
          })
        }
      }

      store.setCells(sheet, updates, 'Remove Duplicates')
    })
  }

  return {
    removedCount,
    remainingCount: uniqueRows.length
  }
}

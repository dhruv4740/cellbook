// Pagination engine: turns a sheet + PageSetup into an ordered list of pages (row/col spans),
// the same way Excel computes page breaks - from the print area (or used range), scaling
// (percent or fit-to-N-pages), margins, paper size and manual row/column breaks.
import { PageSetup, Range, normRange, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../model/types'
import { WorkbookStore, SheetState } from '../model/store'
import { pageGeometry } from './pageGeometry'

export interface PrintPage {
  pageNumber: number // 1-based, honors "First page number"
  blockIndex: number // which disjoint print-area block this page belongs to
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

export interface PrintPlan {
  pages: PrintPage[]
  scalePercent: number // the effective scale actually used (post fit-to-page computation)
  titleRows?: { r1: number; r2: number }
  titleCols?: { c1: number; c2: number }
  /** True when the sheet has nothing to print (empty used range, no print area). */
  empty: boolean
}

function colW(store: WorkbookStore, sheet: SheetState, c: number): number {
  return sheet.layout.hiddenCols[c] ? 0 : sheet.layout.colWidths[c] ?? DEFAULT_COL_WIDTH
}
function rowH(store: WorkbookStore, sheet: SheetState, r: number): number {
  return sheet.layout.hiddenRows[r] ? 0 : sheet.layout.rowHeights[r] ?? DEFAULT_ROW_HEIGHT
}

/** The block(s) of cells to print: the print area (one or more disjoint ranges) if set, else the
 *  sheet's used range as a single block. */
function printBlocks(store: WorkbookStore, sheet: SheetState, ps: PageSetup): Range[] {
  if (ps.printArea.length) return ps.printArea.map(normRange)
  const used = store.usedRange(sheet)
  return used ? [used] : []
}

function sumWidths(store: WorkbookStore, sheet: SheetState, c1: number, c2: number): number {
  let sum = 0
  for (let c = c1; c <= c2; c++) sum += colW(store, sheet, c)
  return sum
}
function sumHeights(store: WorkbookStore, sheet: SheetState, r1: number, r2: number): number {
  let sum = 0
  for (let r = r1; r <= r2; r++) sum += rowH(store, sheet, r)
  return sum
}

/** Resolves the effective print scale: the user's scalePercent, or (when "Fit to" is on) the
 *  largest scale that keeps the block within fitToWidth x fitToHeight pages. 0 in fitToWidth or
 *  fitToHeight means "automatic" (unconstrained) on that axis, matching Excel. */
function resolveScale(store: WorkbookStore, sheet: SheetState, ps: PageSetup, blocks: Range[]): number {
  if (!ps.fitToPage) return Math.max(10, Math.min(400, ps.scalePercent || 100))
  const geo = pageGeometry(ps)
  const titleColsW = ps.printTitleCols ? sumWidths(store, sheet, ps.printTitleCols.c1, ps.printTitleCols.c2) : 0
  const titleRowsH = ps.printTitleRows ? sumHeights(store, sheet, ps.printTitleRows.r1, ps.printTitleRows.r2) : 0
  let contentW = 0, contentH = 0
  for (const b of blocks) {
    contentW = Math.max(contentW, sumWidths(store, sheet, b.c1, b.c2))
    contentH = Math.max(contentH, sumHeights(store, sheet, b.r1, b.r2))
  }
  let scale = 400
  if (ps.fitToWidth > 0) {
    const avail = (geo.printableWidthPx - titleColsW) * ps.fitToWidth
    if (contentW > 0) scale = Math.min(scale, (avail / contentW) * 100)
  }
  if (ps.fitToHeight > 0) {
    const avail = (geo.printableHeightPx - titleRowsH) * ps.fitToHeight
    if (contentH > 0) scale = Math.min(scale, (avail / contentH) * 100)
  }
  // Floor (not round): rounding up could make the scaled content a hair wider/taller than the
  // printable area it was just sized to fit, re-splitting it onto an extra page.
  return Math.max(10, Math.min(400, Math.floor(scale)))
}

export function computePrintPlan(store: WorkbookStore, sheet: SheetState, psIn?: PageSetup): PrintPlan {
  const ps = psIn ?? sheet.layout.pageSetup!
  const blocks = printBlocks(store, sheet, ps)
  if (!blocks.length) return { pages: [], scalePercent: ps.fitToPage ? 100 : ps.scalePercent, empty: true }

  const scale = resolveScale(store, sheet, ps, blocks)
  const factor = scale / 100
  const geo = pageGeometry(ps)
  const titleCols = ps.printTitleCols
  const titleRows = ps.printTitleRows
  const titleColsW = titleCols ? sumWidths(store, sheet, titleCols.c1, titleCols.c2) * factor : 0
  const titleRowsH = titleRows ? sumHeights(store, sheet, titleRows.r1, titleRows.r2) * factor : 0

  const rowBreaks = new Set(ps.rowBreaks)
  const colBreaks = new Set(ps.colBreaks)

  const pages: PrintPage[] = []
  let pageNumber = ps.useFirstPageNumber ? ps.firstPageNumber : 1

  const colRuns = (blockIdx: number, r1: number, r2: number): void => {
    const b = blocks[blockIdx]
    let c = b.c1
    // Skip leading title columns in the content walk - they're rendered on every page already.
    if (titleCols) while (c <= b.c2 && c >= titleCols.c1 && c <= titleCols.c2) c++
    while (c <= b.c2) {
      let width = titleColsW
      const colStart = c
      let colEnd = c
      let first = true
      while (c <= b.c2) {
        const w = colW(store, sheet, c) * factor
        if (!first && width + w > geo.printableWidthPx && width > titleColsW) break
        if (colBreaks.has(c) && !first) break
        width += w
        colEnd = c
        first = false
        c++
      }
      pages.push({ pageNumber: 0, blockIndex: blockIdx, rowStart: r1, rowEnd: r2, colStart, colEnd })
      // A manual break recorded ON colEnd means "break after this column".
      if (colBreaks.has(colEnd + 1)) continue
    }
  }

  blocks.forEach((b, blockIdx) => {
    let r = b.r1
    if (titleRows) while (r <= b.r2 && r >= titleRows.r1 && r <= titleRows.r2) r++
    while (r <= b.r2) {
      let height = titleRowsH
      const rowStart = r
      let rowEnd = r
      let first = true
      while (r <= b.r2) {
        const h = rowH(store, sheet, r) * factor
        if (!first && height + h > geo.printableHeightPx && height > titleRowsH) break
        if (rowBreaks.has(r) && !first) break
        height += h
        rowEnd = r
        first = false
        r++
      }
      colRuns(blockIdx, rowStart, rowEnd)
    }
  })

  // Page order: "over then down" walks all column-pages of the first row-band before moving to
  // the next row-band already matches our nested loop (rows outer). "down then over" (Excel's
  // default) instead finishes every row-page of a column-band first - re-sort accordingly.
  if (ps.pageOrder === 'downThenOver') {
    const colStarts = [...new Set(pages.map((p) => p.colStart))].sort((a, b) => a - b)
    pages.sort((a, b) => {
      if (a.blockIndex !== b.blockIndex) return a.blockIndex - b.blockIndex
      const ca = colStarts.indexOf(a.colStart), cb = colStarts.indexOf(b.colStart)
      if (ca !== cb) return ca - cb
      return a.rowStart - b.rowStart
    })
  }
  pages.forEach((p, i) => (p.pageNumber = pageNumber + i))

  return { pages, scalePercent: scale, titleRows, titleCols, empty: false }
}

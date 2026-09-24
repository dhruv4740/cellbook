// Pure, store-agnostic logic for turning a grid of cells into chart-ready data.
// Kept dependency-free (no HyperFormula, no React) so it's cheap to unit test and reuse from
// the overlay renderer alike.
import { CellValue, Range, isError, normRange } from '../model/types'

/** Minimal read surface a sheet/store must provide; lets tests use a plain object instead of HyperFormula. */
export interface CellReader {
  get(r: number, c: number): CellValue
}

export interface ChartSeries {
  name: string
  values: (number | null)[]
}

export interface ExtractedChartData {
  categories: string[]
  series: ChartSeries[]
}

function numOf(v: CellValue): number | null {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  return null
}

function textOf(v: CellValue): string {
  if (v === null || v === undefined) return ''
  if (isError(v)) return v.error
  return String(v)
}

function isBlank(v: CellValue): boolean {
  return v === null || v === undefined || v === ''
}

/** True if every non-blank cell in the row is text, with at least two of them (a lone label isn't enough evidence of a header). */
function rowIsAllText(cells: CellReader, r: number, c1: number, c2: number): boolean {
  let count = 0
  for (let c = c1; c <= c2; c++) {
    const v = cells.get(r, c)
    if (isBlank(v)) continue
    if (typeof v !== 'string') return false
    count++
  }
  return count >= 2
}

/** True if every non-blank cell in the column is text, with at least two of them. */
function colIsAllText(cells: CellReader, c: number, r1: number, r2: number): boolean {
  let count = 0
  for (let r = r1; r <= r2; r++) {
    const v = cells.get(r, c)
    if (isBlank(v)) continue
    if (typeof v !== 'string') return false
    count++
  }
  return count >= 2
}

/**
 * Extracts categories + series from `range`, matching Excel's default read of a selected range:
 * a text top row becomes series names (or category labels, depending on orientation) and a text
 * left column becomes the other axis's labels, when present.
 */
export function extractChartData(cells: CellReader, range: Range, seriesInColumns: boolean): ExtractedChartData {
  const { r1, c1, r2, c2 } = normRange(range)
  const rows = r2 - r1 + 1
  const cols = c2 - c1 + 1

  const hasColHeader = rows > 1 && rowIsAllText(cells, r1, c1, c2)
  const hasRowHeader = cols > 1 && colIsAllText(cells, c1, r1, r2)

  const dataR1 = r1 + (hasColHeader ? 1 : 0)
  const dataC1 = c1 + (hasRowHeader ? 1 : 0)

  const categories: string[] = []
  const series: ChartSeries[] = []

  if (seriesInColumns) {
    for (let r = dataR1; r <= r2; r++) {
      categories.push(hasRowHeader ? textOf(cells.get(r, c1)) : String(r - r1 + 1))
    }
    for (let c = dataC1; c <= c2; c++) {
      const name = hasColHeader ? textOf(cells.get(r1, c)) : `Series${c - dataC1 + 1}`
      const values: (number | null)[] = []
      for (let r = dataR1; r <= r2; r++) values.push(numOf(cells.get(r, c)))
      series.push({ name, values })
    }
  } else {
    for (let c = dataC1; c <= c2; c++) {
      categories.push(hasColHeader ? textOf(cells.get(r1, c)) : String(c - c1 + 1))
    }
    for (let r = dataR1; r <= r2; r++) {
      const name = hasRowHeader ? textOf(cells.get(r, c1)) : `Series${r - dataR1 + 1}`
      const values: (number | null)[] = []
      for (let c = dataC1; c <= c2; c++) values.push(numOf(cells.get(r, c)))
      series.push({ name, values })
    }
  }
  return { categories, series }
}

/**
 * Picks the orientation Excel defaults to: series run along whichever axis has fewer members,
 * so categories (the longer axis) read along the bottom/left.
 */
export function defaultSeriesInColumns(range: Range): boolean {
  const n = normRange(range)
  const rows = n.r2 - n.r1 + 1
  const cols = n.c2 - n.c1 + 1
  return rows >= cols
}

/**
 * Flood-fills the contiguous block of non-blank cells containing (r,c) — Excel's "current region"
 * (Ctrl+*), used when a chart is inserted from a single active cell rather than a selected range.
 */
export function currentRegion(cells: CellReader, r: number, c: number, maxR: number, maxC: number): Range {
  if (isBlank(cells.get(r, c))) return { r1: r, c1: c, r2: r, c2: c }
  let r1 = r, r2 = r, c1 = c, c2 = c
  let grew = true
  while (grew) {
    grew = false
    while (r1 > 0 && rowHasContent(cells, r1 - 1, c1, c2)) { r1--; grew = true }
    while (r2 < maxR && rowHasContent(cells, r2 + 1, c1, c2)) { r2++; grew = true }
    while (c1 > 0 && colHasContent(cells, c1 - 1, r1, r2)) { c1--; grew = true }
    while (c2 < maxC && colHasContent(cells, c2 + 1, r1, r2)) { c2++; grew = true }
  }
  return { r1, c1, r2, c2 }
}

function rowHasContent(cells: CellReader, r: number, c1: number, c2: number): boolean {
  for (let c = c1; c <= c2; c++) if (!isBlank(cells.get(r, c))) return true
  return false
}
function colHasContent(cells: CellReader, c: number, r1: number, r2: number): boolean {
  for (let r = r1; r <= r2; r++) if (!isBlank(cells.get(r, c))) return true
  return false
}

/** Resolves the range a new chart should plot: the selection if it spans more than one cell, else the current region around the active cell. */
export function resolveChartRange(cells: CellReader, selection: Range, maxR: number, maxC: number): Range {
  const n = normRange(selection)
  if (n.r1 !== n.r2 || n.c1 !== n.c2) return n
  return currentRegion(cells, n.r1, n.c1, maxR, maxC)
}

import { describe, it, expect } from 'vitest'
import { CellValue } from '../model/types'
import { CellReader, extractChartData, defaultSeriesInColumns, currentRegion, resolveChartRange } from './dataExtraction'

function grid(rows: (CellValue)[][]): CellReader {
  return {
    get(r, c) {
      return rows[r]?.[c] ?? null
    }
  }
}

describe('extractChartData', () => {
  it('reads column headers and row labels (series in columns)', () => {
    const cells = grid([
      [null, 'Q1', 'Q2'],
      ['Apples', 5, 7],
      ['Pears', 3, 9]
    ])
    const data = extractChartData(cells, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    expect(data.categories).toEqual(['Apples', 'Pears'])
    expect(data.series).toEqual([
      { name: 'Q1', values: [5, 3] },
      { name: 'Q2', values: [7, 9] }
    ])
  })

  it('reads the same range with series in rows', () => {
    const cells = grid([
      [null, 'Q1', 'Q2'],
      ['Apples', 5, 7],
      ['Pears', 3, 9]
    ])
    const data = extractChartData(cells, { r1: 0, c1: 0, r2: 2, c2: 2 }, false)
    expect(data.categories).toEqual(['Q1', 'Q2'])
    expect(data.series).toEqual([
      { name: 'Apples', values: [5, 7] },
      { name: 'Pears', values: [3, 9] }
    ])
  })

  it('falls back to generic names/labels with no headers', () => {
    const cells = grid([
      [1, 2],
      [3, 4]
    ])
    const data = extractChartData(cells, { r1: 0, c1: 0, r2: 1, c2: 1 }, true)
    expect(data.categories).toEqual(['1', '2'])
    expect(data.series).toEqual([
      { name: 'Series1', values: [1, 3] },
      { name: 'Series2', values: [2, 4] }
    ])
  })

  it('treats blanks and text values as null in numeric series', () => {
    const cells = grid([
      ['', 'A', 'B'],
      ['x', 5, 'n/a'],
      ['y', 'n/a', 9]
    ])
    const data = extractChartData(cells, { r1: 0, c1: 0, r2: 2, c2: 2 }, true)
    expect(data.series[0].values).toEqual([5, null])
    expect(data.series[1].values).toEqual([null, 9])
  })

  it('does not treat a single non-blank corner as a header when the rest of the row is empty', () => {
    // top row only has a value in the corner cell — not "all text", so no header row.
    const cells = grid([
      ['Label', null, null],
      [1, 2, 3]
    ])
    const data = extractChartData(cells, { r1: 0, c1: 0, r2: 1, c2: 2 }, true)
    expect(data.categories).toEqual(['1', '2'])
  })
})

describe('defaultSeriesInColumns', () => {
  it('prefers more categories over more series: tall ranges keep series in columns', () => {
    expect(defaultSeriesInColumns({ r1: 0, c1: 0, r2: 10, c2: 2 })).toBe(true)
  })
  it('wide ranges flip to series in rows', () => {
    expect(defaultSeriesInColumns({ r1: 0, c1: 0, r2: 2, c2: 10 })).toBe(false)
  })
})

describe('currentRegion', () => {
  it('flood-fills the contiguous block around a cell', () => {
    const cells = grid([
      ['A', 'B', null],
      [1, 2, null],
      [3, 4, null],
      [null, null, null]
    ])
    expect(currentRegion(cells, 1, 0, 10, 10)).toEqual({ r1: 0, c1: 0, r2: 2, c2: 1 })
  })

  it('returns a single cell when the anchor cell itself is blank', () => {
    const cells = grid([[null]])
    expect(currentRegion(cells, 0, 0, 10, 10)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 })
  })

  it('stops at gaps: an empty row does not connect two blocks', () => {
    const cells = grid([
      ['A', 1],
      [null, null],
      ['B', 2]
    ])
    expect(currentRegion(cells, 0, 0, 10, 10)).toEqual({ r1: 0, c1: 0, r2: 0, c2: 1 })
  })
})

describe('resolveChartRange', () => {
  it('uses the selection when it spans multiple cells', () => {
    const cells = grid([['A', 'B'], [1, 2]])
    expect(resolveChartRange(cells, { r1: 0, c1: 0, r2: 1, c2: 1 }, 10, 10)).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 })
  })
  it('expands to the current region for a single-cell selection', () => {
    const cells = grid([['A', 'B'], [1, 2]])
    expect(resolveChartRange(cells, { r1: 0, c1: 0, r2: 0, c2: 0 }, 10, 10)).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 })
  })
})

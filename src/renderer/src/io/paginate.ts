// Splits a print range into pages the way Excel does: greedy fill by row height / column width at the
// current scale, manual breaks, repeated title rows/columns, fit-to-pages, and Down-then-over / Over-then-down order.
import { Range } from '../model/types'

export interface PaginateInput {
  range: Range
  /** Screen pixels (96 dpi) of each visible column / row; 0 for hidden ones. */
  colPx: (c: number) => number
  rowPx: (r: number) => number
  /** Printable area of one page in inches (paper minus margins). */
  areaW: number
  areaH: number
  scale: number // percent, ignored when fit is set
  fit?: { w: number; h: number } // pages wide / tall, 0 = automatic
  titleRows?: [number, number]
  titleCols?: [number, number]
  rowBreaks: number[]
  colBreaks: number[]
  order: 'downThenOver' | 'overThenDown'
  firstPageNumber?: number
}

export interface PrintPage {
  number: number
  r1: number
  r2: number
  c1: number
  c2: number
  /** True when the title rows / columns must be drawn above / left of this page's cells. */
  repeatRows: boolean
  repeatCols: boolean
}

export interface Pagination { pages: PrintPage[]; scale: number; pagesWide: number; pagesTall: number; colSpans: [number, number][]; rowSpans: [number, number][] }

const IN = 96

/** Greedy split of lines [from,to] into spans no taller than `avail` (px, unscaled), honouring manual breaks. */
function split(from: number, to: number, size: (i: number) => number, avail: number, breaks: number[], title: [number, number] | undefined): [number, number][] {
  const titleH = title ? sum(title[0], title[1], size) : 0
  const spans: [number, number][] = []
  const breakSet = new Set(breaks)
  let start = from
  while (start <= to) {
    const room = start > (title?.[1] ?? -1) && title ? avail - titleH : avail
    let used = 0
    let end = start - 1
    for (let i = start; i <= to; i++) {
      if (i > start && breakSet.has(i)) break
      const s = size(i)
      if (used + s > room + 1e-6 && i > start) break
      used += s
      end = i
    }
    if (end < start) end = start
    spans.push([start, end])
    start = end + 1
  }
  return spans
}
function sum(a: number, b: number, size: (i: number) => number): number {
  let t = 0
  for (let i = a; i <= b; i++) t += size(i)
  return t
}

function run(inp: PaginateInput, scale: number, useBreaks: boolean): Pagination {
  const { range } = inp
  const availW = (inp.areaW * IN * 100) / scale
  const availH = (inp.areaH * IN * 100) / scale
  const colSpans = split(range.c1, range.c2, inp.colPx, availW, useBreaks ? inp.colBreaks : [], inp.titleCols)
  const rowSpans = split(range.r1, range.r2, inp.rowPx, availH, useBreaks ? inp.rowBreaks : [], inp.titleRows)
  const pages: PrintPage[] = []
  const mk = (ri: number, ci: number): PrintPage => ({
    number: 0, r1: rowSpans[ri][0], r2: rowSpans[ri][1], c1: colSpans[ci][0], c2: colSpans[ci][1],
    repeatRows: !!inp.titleRows && rowSpans[ri][0] > inp.titleRows[1], repeatCols: !!inp.titleCols && colSpans[ci][0] > inp.titleCols[1]
  })
  if (inp.order === 'downThenOver') for (let ci = 0; ci < colSpans.length; ci++) for (let ri = 0; ri < rowSpans.length; ri++) pages.push(mk(ri, ci))
  else for (let ri = 0; ri < rowSpans.length; ri++) for (let ci = 0; ci < colSpans.length; ci++) pages.push(mk(ri, ci))
  const first = inp.firstPageNumber ?? 1
  pages.forEach((p, i) => { p.number = first + i })
  return { pages, scale, pagesWide: colSpans.length, pagesTall: rowSpans.length, colSpans, rowSpans }
}

export function paginate(inp: PaginateInput): Pagination {
  if (!inp.fit) return run(inp, Math.min(400, Math.max(10, inp.scale)), true)
  // Fit to W x T pages: the largest scale (<= 100) whose grid of pages is no bigger; manual breaks are ignored.
  const ok = (s: number): boolean => {
    const p = run(inp, s, false)
    return (inp.fit!.w === 0 || p.pagesWide <= inp.fit!.w) && (inp.fit!.h === 0 || p.pagesTall <= inp.fit!.h)
  }
  if (ok(100)) return run(inp, 100, false)
  let lo = 10, hi = 100
  if (!ok(lo)) return run(inp, lo, false)
  while (hi - lo > 1) {
    const mid = Math.floor((lo + hi) / 2)
    if (ok(mid)) lo = mid
    else hi = mid
  }
  return run(inp, lo, false)
}

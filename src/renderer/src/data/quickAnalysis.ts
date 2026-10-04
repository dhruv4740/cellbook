// Quick Analysis > Totals: the formulas written below (column totals) or to the right (row totals) of a selection.
import { Range, a1, colName, normRange } from '../model/types'

export type TotalsFn = 'SUM' | 'AVERAGE' | 'COUNT'
export type TotalsAxis = 'col' | 'row'

/** Cells to write for one Totals button. Lines with no numbers are skipped, so header text columns get no total. */
export function planTotals(sel: Range, fn: TotalsFn, axis: TotalsAxis, isNum: (r: number, c: number) => boolean): { r: number; c: number; raw: string }[] {
  const { r1, c1, r2, c2 } = normRange(sel)
  const out: { r: number; c: number; raw: string }[] = []
  if (axis === 'col') {
    for (let c = c1; c <= c2; c++) {
      let any = false
      for (let r = r1; r <= r2 && !any; r++) any = isNum(r, c)
      if (any) out.push({ r: r2 + 1, c, raw: `=${fn}(${colName(c)}${r1 + 1}:${colName(c)}${r2 + 1})` })
    }
  } else {
    for (let r = r1; r <= r2; r++) {
      let any = false
      for (let c = c1; c <= c2 && !any; c++) any = isNum(r, c)
      if (any) out.push({ r, c: c2 + 1, raw: `=${fn}(${a1(r, c1)}:${a1(r, c2)})` })
    }
  }
  return out
}

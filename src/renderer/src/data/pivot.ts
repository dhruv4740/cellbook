// Insert > PivotTable: summarises a list by row/column fields with Sum/Count/Average/Max/Min, laid
// out like Excel's default (compact) PivotTable with grand totals. The definition is kept on the
// destination sheet so Data > Refresh All can rebuild it after the source data changes.
import type { WorkbookStore, SheetState } from '../model/store'
import { CellStyle, Range } from '../model/types'

export type PivotFunc = 'sum' | 'count' | 'average' | 'max' | 'min'
export const PIVOT_FUNC_LABEL: Record<PivotFunc, string> = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min' }

export interface PivotDef {
  id: string
  sourceSheet: string
  source: Range // includes the header row
  anchor: { r: number; c: number } // top-left of the output on the destination sheet
  rows: number[] // source columns (absolute)
  cols: number[] // at most one column field
  values: { col: number; func: PivotFunc }[]
  /** Size of the last output, so a refresh clears exactly what it wrote. */
  lastSize?: { rows: number; cols: number }
}

type Cell = { raw: string; style?: CellStyle }

const HEADER: CellStyle = { bold: true, fill: '#DDEBF7', borders: { bottom: { style: 'thin', color: '#9BC2E6' } } }
const TOTAL: CellStyle = { bold: true, fill: '#DDEBF7', borders: { top: { style: 'thin', color: '#9BC2E6' } } }
const SUB: CellStyle = { bold: true }

function agg(func: PivotFunc, vals: unknown[]): number | '' {
  if (func === 'count') return vals.filter((v) => v !== null && v !== undefined && v !== '').length
  const nums = vals.filter((v): v is number => typeof v === 'number')
  if (!nums.length) return func === 'sum' ? 0 : ''
  switch (func) {
    case 'sum': return nums.reduce((a, b) => a + b, 0)
    case 'average': return nums.reduce((a, b) => a + b, 0) / nums.length
    case 'max': return Math.max(...nums)
    case 'min': return Math.min(...nums)
  }
}

function cmpKey(a: string, b: string): number {
  if (a === '(blank)') return 1
  if (b === '(blank)') return -1
  const na = Number(a), nb = Number(b)
  if (a !== '' && b !== '' && !isNaN(na) && !isNaN(nb)) return na - nb
  return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true })
}

/** Builds the output grid (rows of cells) for a pivot definition. */
export function computePivot(store: WorkbookStore, def: PivotDef): Cell[][] | string {
  const src = store.sheets.find((s) => s.name === def.sourceSheet)
  if (!src) return `The source sheet '${def.sourceSheet}' no longer exists.`
  const { r1, r2 } = def.source
  if (r2 <= r1) return 'The source range needs a header row and at least one row of data.'
  if (!def.values.length && !def.rows.length) return 'Choose at least one field for Rows or Values.'
  const head = (c: number) => store.getDisplay(src, r1, c).text || `Column${c + 1}`
  const label = (r: number, c: number) => store.getDisplay(src, r, c).text || '(blank)'
  const values = def.values.length ? def.values : [{ col: def.rows[0], func: 'count' as PivotFunc }]
  const vName = (v: { col: number; func: PivotFunc }) => `${PIVOT_FUNC_LABEL[v.func]} of ${head(v.col)}`

  // Collect records.
  const recs: { rk: string[]; ck: string; vals: unknown[] }[] = []
  for (let r = r1 + 1; r <= r2; r++) {
    if (src.layout.autoFilter?.filteredOutRows.includes(r)) continue
    let empty = true
    for (let c = def.source.c1; c <= def.source.c2; c++) if (store.getRaw(src, r, c) !== '') { empty = false; break }
    if (empty) continue
    recs.push({ rk: def.rows.map((c) => label(r, c)), ck: def.cols.length ? label(r, def.cols[0]) : '', vals: values.map((v) => store.getValue(src, r, v.col)) })
  }
  const colKeys = def.cols.length ? [...new Set(recs.map((x) => x.ck))].sort(cmpKey) : ['']
  const aggFor = (subset: typeof recs, vi: number): number | '' => agg(values[vi].func, subset.map((x) => x.vals[vi]))

  const out: Cell[][] = []
  const withCols = def.cols.length > 0
  // Header rows
  if (withCols) {
    out.push([{ raw: vName(values[0]), style: HEADER }, { raw: 'Column Labels', style: HEADER }, ...colKeys.slice(1).map(() => ({ raw: '', style: HEADER })), { raw: '', style: HEADER }])
    out.push([{ raw: 'Row Labels', style: HEADER }, ...colKeys.map((k) => ({ raw: k, style: HEADER })), { raw: 'Grand Total', style: HEADER }])
  } else {
    out.push([{ raw: def.rows.length ? 'Row Labels' : '', style: HEADER }, ...values.map((v) => ({ raw: vName(v), style: HEADER }))])
  }
  const num = (x: number | '') => (x === '' ? '' : String(x))
  const valueCells = (subset: typeof recs, style?: CellStyle): Cell[] => withCols
    ? [...colKeys.map((k) => ({ raw: num(subset.some((x) => x.ck === k) ? aggFor(subset.filter((x) => x.ck === k), 0) : ''), style })), { raw: num(aggFor(subset, 0)), style }]
    : values.map((_, vi) => ({ raw: num(aggFor(subset, vi)), style }))

  if (def.rows.length) {
    const outer = [...new Set(recs.map((x) => x.rk[0]))].sort(cmpKey)
    for (const o of outer) {
      const sub = recs.filter((x) => x.rk[0] === o)
      out.push([{ raw: o, style: def.rows.length > 1 ? SUB : undefined }, ...valueCells(sub, def.rows.length > 1 ? SUB : undefined)])
      if (def.rows.length > 1) {
        const inner = [...new Set(sub.map((x) => x.rk[1]))].sort(cmpKey)
        for (const i of inner) out.push([{ raw: i, style: { indent: 1 } }, ...valueCells(sub.filter((x) => x.rk[1] === i))])
      }
    }
  }
  out.push([{ raw: 'Grand Total', style: TOTAL }, ...valueCells(recs, TOTAL)])
  // Text labels that look like numbers/formulas stay text.
  for (const row of out) row[0] = { ...row[0], raw: /^[=+\-@]|^\d/.test(row[0].raw) ? "'" + row[0].raw : row[0].raw }
  if (withCols) out[1] = out[1].map((c) => (/^[=+\-@]|^\d/.test(c.raw) ? { ...c, raw: "'" + c.raw } : c))
  return out
}

/** Writes (or rewrites) a pivot's output on `dest`. Returns an error message or null. */
export function renderPivot(store: WorkbookStore, dest: SheetState, def: PivotDef): string | null {
  const grid = computePivot(store, def)
  if (typeof grid === 'string') return grid
  const width = Math.max(...grid.map((r) => r.length))
  const { r, c } = def.anchor
  store.transact('PivotTable', () => {
    const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
    if (def.lastSize) for (let i = 0; i < def.lastSize.rows; i++) for (let j = 0; j < def.lastSize.cols; j++) cells.push({ r: r + i, c: c + j, raw: '', style: undefined })
    const idx = new Map(cells.map((x, k) => [`${x.r},${x.c}`, k]))
    grid.forEach((row, i) => row.forEach((cell, j) => {
      const k = idx.get(`${r + i},${c + j}`)
      const v = { r: r + i, c: c + j, raw: cell.raw, style: cell.style }
      if (k !== undefined) cells[k] = v; else cells.push(v)
    }))
    store.setCells(dest, cells, 'PivotTable')
    const next: PivotDef = { ...def, lastSize: { rows: grid.length, cols: width } }
    store.updateLayout(dest, 'PivotTable', (l) => {
      l.pivots = [...(l.pivots ?? []).filter((p) => p.id !== def.id), next]
    })
    // Size label columns like Excel's autofit on create.
    if (!def.lastSize) {
      const w0 = Math.max(90, ...grid.map((row) => (row[0]?.raw.length ?? 0) * 7 + 12))
      store.setColWidth(dest, [c], Math.min(300, w0))
      for (let j = 1; j < width; j++) store.setColWidth(dest, [c + j], Math.min(200, Math.max(80, ...grid.map((row) => (row[j]?.raw.length ?? 0) * 7 + 12))))
    }
  })
  return null
}

/** Data > Refresh All: rebuilds every PivotTable in the workbook from its current source data. */
export function refreshAllPivots(store: WorkbookStore): string[] {
  const errs: string[] = []
  for (const sh of store.sheets) for (const p of sh.layout.pivots ?? []) {
    const e = renderPivot(store, sh, p)
    if (e) errs.push(e)
  }
  return errs
}

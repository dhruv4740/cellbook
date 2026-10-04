// Data > Forecast Sheet: builds a new sheet with the timeline, the history, a forecast column and
// lower/upper confidence bounds as live formulas. The engine has no FORECAST.ETS, so the forecast
// is a linear trend (FORECAST.LINEAR) with a regression prediction interval (STEYX + T.INV.2T).
import type { WorkbookStore, SheetState } from '../model/store'
import { Range, normRange } from '../model/types'
import { insertChart } from '../charts/chartsStore'
import { createChartSpec, DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT } from '../charts/gallery'

export interface ForecastOptions {
  chart: 'line' | 'column'
  end: number // last timeline value to forecast
  confidence: number | null // 0..1, null = no bounds
}

export interface ForecastSource { xs: number[]; ys: number[]; xHead: string; yHead: string }

/** Timeline step: the average gap of the history. */
export function forecastStep(xs: number[]): number {
  return (xs[xs.length - 1] - xs[0]) / (xs.length - 1)
}

/** Parses the Forecast End box: a number, or a date such as 2026-12-31 or 12/31/2026 (-> Excel serial). */
export function parseForecastEnd(text: string): number | null {
  const t = text.trim()
  if (!t) return null
  if (!isNaN(Number(t))) return Number(t)
  const d = Date.parse(t + ' UTC')
  return isNaN(d) ? null : Math.round(d / 86400000) + 25569
}

/** Raw cell text for the new sheet, row-major (row 0 = headers). Parameters live in G1:H2. */
export function buildForecastGrid(src: ForecastSource, opts: ForecastOptions): string[][] {
  const n = src.xs.length
  const step = forecastStep(src.xs)
  const future = Math.max(0, Math.floor((opts.end - src.xs[n - 1]) / step + 1e-9))
  const xr = `$A$2:$A$${n + 1}`
  const yr = `$B$2:$B$${n + 1}`
  const bounds = opts.confidence !== null
  const h = src.yHead
  const rows: string[][] = [[src.xHead, h, `Forecast(${h})`, ...(bounds ? [`Lower Confidence Bound(${h})`, `Upper Confidence Bound(${h})`] : [])]]
  for (let i = 0; i < n + future; i++) {
    const r = i + 2
    const row: string[] = ['', '', '']
    if (i < n) {
      row[0] = String(src.xs[i]); row[1] = String(src.ys[i])
      if (i === n - 1) row[2] = `=B${r}` // joins the forecast line to the history
    } else {
      row[0] = `=A${r - 1}+$H$2`
      row[2] = `=FORECAST.LINEAR(A${r},${yr},${xr})`
    }
    if (bounds) {
      if (i < n - 1) row.push('', '')
      else if (i === n - 1) row.push(`=C${r}`, `=C${r}`)
      else {
        const m = `T.INV.2T(1-$H$1,COUNT(${xr})-2)*STEYX(${yr},${xr})*SQRT(1+1/COUNT(${xr})+(A${r}-AVERAGE(${xr}))^2/DEVSQ(${xr}))`
        row.push(`=C${r}-${m}`, `=C${r}+${m}`)
      }
    }
    rows.push(row)
  }
  for (const r of rows) while (r.length < 8) r.push('')
  rows[0][6] = 'Confidence Interval'; rows[0][7] = bounds ? String(opts.confidence) : ''
  rows[1][6] = 'Timeline Step'; rows[1][7] = String(step)
  return rows
}

/** Reads the selection as timeline + values (two columns, or two rows). Returns an error message if it is not usable. */
export function readForecastSource(store: WorkbookStore, sheet: SheetState, rg: Range): ForecastSource | string {
  const n = normRange(rg)
  const byRows = n.r2 - n.r1 === 1 && n.c2 - n.c1 > 1
  const len = byRows ? n.c2 - n.c1 + 1 : n.r2 - n.r1 + 1
  if (!byRows && n.c2 - n.c1 !== 1) return 'Select two columns (or two rows): a timeline and the values.'
  const cell = (k: number, w: 0 | 1) => (byRows ? store.getValue(sheet, n.r1 + w, n.c1 + k) : store.getValue(sheet, n.r1 + k, n.c1 + w))
  const hasHead = typeof cell(0, 0) !== 'number'
  const xs: number[] = [], ys: number[] = []
  for (let k = hasHead ? 1 : 0; k < len; k++) {
    const x = cell(k, 0), y = cell(k, 1)
    if (typeof x !== 'number' || typeof y !== 'number') return 'The timeline and the values must be numbers or dates.'
    xs.push(x); ys.push(y)
  }
  if (xs.length < 3) return 'Select at least three data points.'
  for (let i = 1; i < xs.length; i++) if (xs[i] <= xs[i - 1]) return 'The timeline must be in ascending order.'
  const head = (w: 0 | 1, d: string) => (hasHead ? store.getDisplay(sheet, byRows ? n.r1 + w : n.r1, byRows ? n.c1 : n.c1 + w).text : d)
  return { xs, ys, xHead: head(0, 'Timeline'), yHead: head(1, 'Values') }
}

/** Adds the forecast sheet (with its chart) after the active one as one action. Returns an error message, or null. */
export function createForecastSheet(store: WorkbookStore, src: ForecastSource, opts: ForecastOptions): string | null {
  if (!(opts.end > src.xs[src.xs.length - 1])) return 'Forecast End must be after the last point of the timeline.'
  const grid = buildForecastGrid(src, opts)
  if (grid.length > 5000) return 'That forecast would be too long; choose an earlier Forecast End.'
  store.transact('Forecast Sheet', () => {
    store.addSheet()
    const sh = store.active
    grid.forEach((row, r) => row.forEach((v, c) => { if (v !== '') store.setRaw(sh, r, c, v) }))
    const spec = createChartSpec(
      opts.chart === 'line' ? { type: 'line', subtype: 'line', label: 'Line' } : { type: 'bar', subtype: 'clusteredColumn', label: 'Clustered Column' },
      { r1: 0, c1: 0, r2: grid.length - 1, c2: opts.confidence !== null ? 4 : 2 },
      { r: 1, c: 9, dx: 0, dy: 0 }
    )
    spec.seriesInColumns = true
    insertChart(store, sh, spec)
  })
  return null
}

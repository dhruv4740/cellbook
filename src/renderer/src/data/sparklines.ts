// Insert > Sparklines: a mini chart in a single cell, drawn from a row or column of numbers.
import { store, SheetState } from '../model/store'
import { parseA1, SparklineDef } from '../model/types'

export const SPARK_COLOR = '#376092'

/** Numbers from the source range in order (blank / text become null so the chart leaves a gap). */
export function sparkValues(sheet: SheetState, s: SparklineDef): (number | null)[] {
  const out: (number | null)[] = []
  for (let r = s.src.r1; r <= s.src.r2; r++)
    for (let c = s.src.c1; c <= s.src.c2; c++) {
      const v = store.getValue(sheet, r, c)
      out.push(typeof v === 'number' && Number.isFinite(v) ? v : null)
    }
  return out
}

export function drawSparkline(ctx: CanvasRenderingContext2D, vals: (number | null)[], type: SparklineDef['type'], color: string, x: number, y: number, w: number, h: number): void {
  const nums = vals.filter((v): v is number => v !== null)
  if (!nums.length || w < 6 || h < 6) return
  const px = 4, py = 4
  const iw = w - 2 * px, ih = h - 2 * py
  let min = Math.min(...nums), max = Math.max(...nums)
  if (type !== 'winloss') { if (type === 'column') { min = Math.min(min, 0); max = Math.max(max, 0) } }
  if (min === max) { min -= 1; max += 1 }
  const yOf = (v: number) => y + py + ih - ((v - min) / (max - min)) * ih
  const n = vals.length
  ctx.save()
  ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip()
  if (type === 'line') {
    ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.beginPath()
    let pen = false
    vals.forEach((v, i) => {
      if (v === null) { pen = false; return }
      const X = x + px + (n === 1 ? iw / 2 : (i / (n - 1)) * iw), Y = yOf(v)
      if (pen) ctx.lineTo(X, Y); else ctx.moveTo(X, Y)
      pen = true
    })
    ctx.stroke()
  } else {
    const slot = iw / n, bw = Math.max(1, slot * 0.7)
    const zeroY = type === 'column' ? yOf(0) : y + h / 2
    vals.forEach((v, i) => {
      if (v === null) return
      const bx = x + px + i * slot + (slot - bw) / 2
      ctx.fillStyle = v < 0 ? '#d00000' : color
      if (type === 'column') {
        const Y = yOf(v)
        ctx.fillRect(bx, Math.min(Y, zeroY), bw, Math.max(1, Math.abs(Y - zeroY)))
      } else if (v !== 0) {
        const half = ih / 2 - 1
        ctx.fillRect(bx, v > 0 ? zeroY - half : zeroY + 1, bw, half)
      }
    })
  }
  ctx.restore()
}

/** Adds one sparkline per row (or column) of `data` into the cells of `location`. Returns an error message, or null. */
export function addSparklines(sheet: SheetState, type: SparklineDef['type'], dataRef: string, locRef: string): string | null {
  const data = parseA1(dataRef.replace(/\$/g, '').replace(/^.*!/, '').trim())
  const loc = parseA1(locRef.replace(/\$/g, '').replace(/^.*!/, '').trim())
  if (!data) return 'The Data Range reference isn’t valid.'
  if (!loc) return 'The Location Range reference isn’t valid.'
  const dr = data.r2 - data.r1 + 1, dc = data.c2 - data.c1 + 1
  const lr = loc.r2 - loc.r1 + 1, lc = loc.c2 - loc.c1 + 1
  const cells = lr * lc
  // Rows of data map to a vertical location range, columns to a horizontal one (Excel's rule); a single row/col
  // of data goes into a single cell.
  const byRow = cells > 1 && lc === 1 && dr === lr ? true : cells > 1 && lr === 1 && dc === lc ? false : cells === 1 ? (dr === 1 ? true : dc === 1 ? false : null) : null
  if (cells === 1 && dr > 1 && dc > 1) return 'The Location Range must match the Data Range: one cell per row or column of data.'
  if (byRow === null) return 'The Location Range doesn’t match the Data Range: it needs one cell for each row (or column) of data.'
  const defs: SparklineDef[] = []
  for (let i = 0; i < cells; i++) {
    const r = loc.r1 + (lr > 1 ? i : 0), c = loc.c1 + (lc > 1 ? i : 0)
    if (cells === 1) defs.push({ r, c, type, src: { ...data }, color: SPARK_COLOR })
    else if (byRow) defs.push({ r, c, type, src: { r1: data.r1 + i, r2: data.r1 + i, c1: data.c1, c2: data.c2 }, color: SPARK_COLOR })
    else defs.push({ r, c, type, src: { r1: data.r1, r2: data.r2, c1: data.c1 + i, c2: data.c1 + i }, color: SPARK_COLOR })
  }
  store.updateLayout(sheet, 'Insert Sparklines', (l) => {
    const at = new Set(defs.map((d) => d.r + ',' + d.c))
    l.sparklines = [...(l.sparklines ?? []).filter((s) => !at.has(s.r + ',' + s.c)), ...defs]
  })
  return null
}

export function clearSparklines(sheet: SheetState, ranges: { r1: number; c1: number; r2: number; c2: number }[]): void {
  const inSel = (r: number, c: number) => ranges.some((g) => r >= Math.min(g.r1, g.r2) && r <= Math.max(g.r1, g.r2) && c >= Math.min(g.c1, g.c2) && c <= Math.max(g.c1, g.c2))
  if (!(sheet.layout.sparklines ?? []).some((s) => inSel(s.r, s.c))) return
  store.updateLayout(sheet, 'Clear Sparklines', (l) => { l.sparklines = (l.sparklines ?? []).filter((s) => !inSel(s.r, s.c)) })
}

// Home > Copy > Copy as Picture: draws the selected range to a PNG (values as displayed, fills, borders, optional
// gridlines) and puts it on the system clipboard through the main process.
import { store, SheetState, WorkbookStore } from '../model/store'
import { CellStyle, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, Range, normRange, normalFont } from '../model/types'
import { showToast } from '../chrome/ui'

export interface PicCell { x: number; y: number; w: number; h: number; text: string; style?: CellStyle; align: 'left' | 'right' | 'center'; color?: string }
export interface PicPlan { width: number; height: number; cells: PicCell[]; grid: { x: number; y: number }[][] }

const MAX_CELLS = 100000
const MAX_PX = 8000

/** Cell rectangles (px at 100%, hidden rows and columns collapsed) for the range; null when it is too big to draw. */
export function planPicture(sheet: SheetState, range: Range, gridlines: boolean, st: WorkbookStore = store): PicPlan | null {
  const rg = normRange(range)
  const xs: number[] = [0], ys: number[] = [0]
  for (let c = rg.c1; c <= rg.c2; c++) xs.push(xs[xs.length - 1] + (sheet.layout.hiddenCols[c] ? 0 : sheet.layout.colWidths[c] ?? DEFAULT_COL_WIDTH))
  for (let r = rg.r1; r <= rg.r2; r++) ys.push(ys[ys.length - 1] + (sheet.layout.hiddenRows[r] ? 0 : sheet.layout.rowHeights[r] ?? DEFAULT_ROW_HEIGHT))
  const width = xs[xs.length - 1], height = ys[ys.length - 1]
  if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > MAX_CELLS || width > MAX_PX || height > MAX_PX || width <= 0 || height <= 0) return null
  const cells: PicCell[] = []
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (sheet.layout.hiddenRows[r]) continue
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (sheet.layout.hiddenCols[c]) continue
      const m = st.mergeAt(sheet, r, c)
      if (m && (r !== Math.max(m.r1, rg.r1) || c !== Math.max(m.c1, rg.c1))) continue // only a merge's top-left cell is drawn
      const r2 = Math.min(m?.r2 ?? r, rg.r2), c2 = Math.min(m?.c2 ?? c, rg.c2)
      const style = st.getStyle(sheet, m?.r1 ?? r, m?.c1 ?? c)
      const d = st.getDisplay(sheet, m?.r1 ?? r, m?.c1 ?? c)
      const x = xs[c - rg.c1], y = ys[r - rg.r1]
      const a = style?.hAlign
      cells.push({ x, y, w: xs[c2 - rg.c1 + 1] - x, h: ys[r2 - rg.r1 + 1] - y, text: d.text, style, align: a === 'left' || a === 'right' || a === 'center' ? a : d.align ?? 'left', color: d.color })
    }
  }
  return { width, height, cells, grid: gridlines ? [xs.map((x) => ({ x, y: 0 })), ys.map((y) => ({ x: 0, y }))] : [[], []] }
}

function font(s: CellStyle | undefined): string {
  return `${s?.italic ? 'italic ' : ''}${s?.bold ? 'bold ' : ''}${Math.round(s?.fontSize ?? normalFont.size)}px "${s?.fontFamily ?? normalFont.name}", "Carlito", sans-serif`
}

export function renderPicture(plan: PicPlan, scale = 2): HTMLCanvasElement {
  const cv = document.createElement('canvas')
  cv.width = Math.ceil(plan.width * scale)
  cv.height = Math.ceil(plan.height * scale)
  const ctx = cv.getContext('2d')!
  ctx.scale(scale, scale)
  ctx.fillStyle = '#fff'
  ctx.fillRect(0, 0, plan.width, plan.height)
  for (const c of plan.cells) if (c.style?.fill) { ctx.fillStyle = c.style.fill; ctx.fillRect(c.x, c.y, c.w, c.h) }
  ctx.strokeStyle = '#d4d4d4'
  ctx.lineWidth = 1
  ctx.beginPath()
  for (const p of plan.grid[0]) { ctx.moveTo(Math.round(p.x) + 0.5, 0); ctx.lineTo(Math.round(p.x) + 0.5, plan.height) }
  for (const p of plan.grid[1]) { ctx.moveTo(0, Math.round(p.y) + 0.5); ctx.lineTo(plan.width, Math.round(p.y) + 0.5) }
  ctx.stroke()
  for (const c of plan.cells) {
    if (c.text) {
      ctx.save()
      ctx.beginPath()
      ctx.rect(c.x, c.y, c.w, c.h)
      ctx.clip()
      ctx.font = font(c.style)
      ctx.fillStyle = c.color ?? c.style?.color ?? '#000'
      ctx.textBaseline = 'middle'
      ctx.textAlign = c.align
      const px = c.align === 'left' ? c.x + 3 : c.align === 'right' ? c.x + c.w - 3 : c.x + c.w / 2
      const py = c.style?.vAlign === 'top' ? c.y + 8 : c.style?.vAlign === 'middle' ? c.y + c.h / 2 : c.y + c.h - 8
      ctx.fillText(c.text, px, py)
      ctx.restore()
    }
    const b = c.style?.borders
    if (b) {
      for (const [side, bs] of Object.entries(b)) {
        if (!bs) continue
        ctx.strokeStyle = bs.color ?? '#000'
        ctx.lineWidth = bs.style === 'thick' ? 3 : bs.style === 'medium' ? 2 : 1
        ctx.beginPath()
        if (side === 'top') { ctx.moveTo(c.x, c.y); ctx.lineTo(c.x + c.w, c.y) }
        else if (side === 'bottom') { ctx.moveTo(c.x, c.y + c.h); ctx.lineTo(c.x + c.w, c.y + c.h) }
        else if (side === 'left') { ctx.moveTo(c.x, c.y); ctx.lineTo(c.x, c.y + c.h) }
        else { ctx.moveTo(c.x + c.w, c.y); ctx.lineTo(c.x + c.w, c.y + c.h) }
        ctx.stroke()
      }
    }
  }
  return cv
}

/** Copies the active selection's last range as a PNG; gridlines follow the sheet's View > Gridlines setting. */
export async function copyAsPicture(): Promise<void> {
  const sheet = store.active
  const sel = sheet.selection.ranges
  const plan = planPicture(sheet, sel[sel.length - 1], sheet.layout.showGridlines)
  if (!plan) { showToast('That range is too large to copy as a picture.'); return }
  const png = renderPicture(plan).toDataURL('image/png')
  if (!window.api?.clipboardWriteImage) { showToast('Copy as Picture needs the desktop app.'); return }
  await window.api.clipboardWriteImage(png)
  showToast('Copied the range as a picture.')
}

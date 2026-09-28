// The outline gutter: level buttons (1 2 3 ...), group brackets and [-]/[+] buttons drawn left of the row
// headers and above the column headers, like Excel. Pure drawing/hit-testing over a geometry snapshot so
// Grid.tsx only supplies the screen positions of rows and columns.
import { Axis, OutlineRun } from '../model/outline'

export const OL = 13 // gutter width/height per group level (px, not scaled by zoom)
const BOX = 11

export interface GutterGeo {
  gutterL: number
  gutterT: number
  headerH: number
  rowHeaderW: number
  showHead: boolean
  size: { w: number; h: number }
  rowRuns: OutlineRun[]
  colRuns: OutlineRun[]
  rowDepth: number
  colDepth: number
  /** Screen y of the top edge of row i (i may be one past the last row of a run). */
  rowY: (i: number) => number
  colX: (i: number) => number
  summaryBelow: boolean
  summaryRight: boolean
}

export type GutterHit = { kind: 'level'; axis: Axis; n: number } | { kind: 'run'; axis: Axis; run: OutlineRun }

type Box = { x: number; y: number; w: number; h: number }

export function gutterSizes(rowDepth: number, colDepth: number, showHead: boolean, hidden: boolean): { gutterL: number; gutterT: number } {
  if (hidden || (rowDepth === 0 && colDepth === 0)) return { gutterL: 0, gutterT: 0 }
  const gutterL = rowDepth > 0 ? OL * (rowDepth + 1) + 3 : 17
  const gutterT = colDepth > 0 ? OL * (colDepth + 1) + 3 + (rowDepth > 0 && !showHead ? 15 : 0) : !showHead ? 17 : 0
  return { gutterL, gutterT }
}

function levelButton(g: GutterGeo, axis: Axis, n: number): Box {
  if (axis === 'row') {
    const y = g.showHead ? g.gutterT + Math.max(0, (g.headerH - g.gutterT - BOX - 1) / 2) : g.gutterT - 15
    return { x: 2 + OL * (n - 1), y: Math.max(1, y), w: BOX + 1, h: BOX + 1 }
  }
  return { x: 2, y: 2 + OL * (n - 1), w: BOX + 1, h: BOX + 1 }
}

function runBox(g: GutterGeo, axis: Axis, run: OutlineRun): Box | null {
  if (run.summary < 0) return null
  if (axis === 'row') {
    const top = g.rowY(run.summary), h = g.rowY(run.summary + 1) - top
    if (h <= 0.5) return null
    const y = top + h / 2 - BOX / 2
    if (y + BOX < g.headerH || y > g.size.h) return null
    return { x: 2 + OL * (run.level - 1), y, w: BOX, h: BOX }
  }
  const left = g.colX(run.summary), w = g.colX(run.summary + 1) - left
  if (w <= 0.5) return null
  const x = left + w / 2 - BOX / 2
  if (x + BOX < g.rowHeaderW || x > g.size.w) return null
  return { x, y: 2 + OL * (run.level - 1), w: BOX, h: BOX }
}

export function hitGutter(g: GutterGeo, x: number, y: number): GutterHit | null {
  const inside = (b: Box | null): boolean => !!b && x >= b.x && x <= b.x + b.w && y >= b.y && y <= b.y + b.h
  if (g.rowDepth > 0) for (let n = 1; n <= g.rowDepth + 1; n++) if (inside(levelButton(g, 'row', n))) return { kind: 'level', axis: 'row', n }
  if (g.colDepth > 0) for (let n = 1; n <= g.colDepth + 1; n++) if (inside(levelButton(g, 'col', n))) return { kind: 'level', axis: 'col', n }
  if (x < g.gutterL && y >= g.headerH) for (const run of g.rowRuns) if (inside(runBox(g, 'row', run))) return { kind: 'run', axis: 'row', run }
  if (y < g.gutterT && x >= g.rowHeaderW) for (const run of g.colRuns) if (inside(runBox(g, 'col', run))) return { kind: 'run', axis: 'col', run }
  return null
}

export function drawGutter(ctx: CanvasRenderingContext2D, g: GutterGeo): void {
  if (g.gutterL === 0 && g.gutterT === 0) return
  const { gutterL, gutterT, headerH, rowHeaderW, size } = g
  ctx.save()
  ctx.fillStyle = '#f5f5f5'
  ctx.fillRect(0, 0, gutterL, size.h)
  ctx.fillRect(0, 0, size.w, gutterT)
  ctx.font = '10px "Segoe UI", sans-serif'
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle'

  const button = (b: Box, label: string): void => {
    ctx.fillStyle = '#fff'; ctx.fillRect(b.x, b.y, b.w, b.h)
    ctx.strokeStyle = '#8a8a8a'; ctx.lineWidth = 1; ctx.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1)
    ctx.fillStyle = '#1f1f1f'; ctx.fillText(label, b.x + b.w / 2, b.y + b.h / 2 + 0.5)
  }
  const sign = (b: Box, collapsed: boolean): void => {
    ctx.fillStyle = '#fff'; ctx.fillRect(b.x, b.y, b.w, b.h)
    ctx.strokeStyle = '#8a8a8a'; ctx.lineWidth = 1; ctx.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1)
    ctx.strokeStyle = '#1f1f1f'
    ctx.beginPath()
    ctx.moveTo(b.x + 3, b.y + b.h / 2); ctx.lineTo(b.x + b.w - 3, b.y + b.h / 2)
    if (collapsed) { ctx.moveTo(b.x + b.w / 2, b.y + 3); ctx.lineTo(b.x + b.w / 2, b.y + b.h - 3) }
    ctx.stroke()
  }

  // Row brackets (clipped to the gutter below the headers).
  if (g.rowDepth > 0) {
    ctx.save(); ctx.beginPath(); ctx.rect(0, headerH, gutterL, size.h - headerH); ctx.clip()
    ctx.strokeStyle = '#1f1f1f'; ctx.lineWidth = 1
    for (const run of g.rowRuns) {
      const bx = Math.round(2 + OL * (run.level - 1) + BOX / 2) + 0.5
      if (!run.collapsed) {
        const y1 = g.rowY(run.start), y2 = g.rowY(run.end + 1)
        if (y2 >= headerH && y1 <= size.h && y2 > y1) {
          ctx.beginPath()
          ctx.moveTo(bx, y1 + 2); ctx.lineTo(bx, y2 - (g.summaryBelow ? 3 : 0))
          const ty = Math.round(g.summaryBelow ? y2 - 3 : y1 + 2) + 0.5
          ctx.moveTo(bx, ty); ctx.lineTo(bx + 5, ty)
          ctx.stroke()
        }
      }
      const b = runBox(g, 'row', run)
      if (b) sign(b, run.collapsed)
    }
    ctx.restore()
  }
  // Column brackets.
  if (g.colDepth > 0) {
    ctx.save(); ctx.beginPath(); ctx.rect(rowHeaderW, 0, size.w - rowHeaderW, gutterT); ctx.clip()
    ctx.strokeStyle = '#1f1f1f'; ctx.lineWidth = 1
    for (const run of g.colRuns) {
      const by = Math.round(2 + OL * (run.level - 1) + BOX / 2) + 0.5
      if (!run.collapsed) {
        const x1 = g.colX(run.start), x2 = g.colX(run.end + 1)
        if (x2 >= rowHeaderW && x1 <= size.w && x2 > x1) {
          ctx.beginPath()
          ctx.moveTo(x1 + 2, by); ctx.lineTo(x2 - (g.summaryRight ? 3 : 0), by)
          const tx = Math.round(g.summaryRight ? x2 - 3 : x1 + 2) + 0.5
          ctx.moveTo(tx, by); ctx.lineTo(tx, by + 5)
          ctx.stroke()
        }
      }
      const b = runBox(g, 'col', run)
      if (b) sign(b, run.collapsed)
    }
    ctx.restore()
  }
  // Level buttons in the corner.
  if (g.rowDepth > 0) for (let n = 1; n <= g.rowDepth + 1; n++) button(levelButton(g, 'row', n), String(n))
  if (g.colDepth > 0) for (let n = 1; n <= g.colDepth + 1; n++) button(levelButton(g, 'col', n), String(n))
  ctx.restore()
}

// Draw tab: freehand ink strokes are ShapeDefs of kind 'ink' (points in the shape's own pixel box), so move / resize /
// delete / arrange / Selection Pane / hide all work unchanged. Pure geometry here so it unit-tests without the store.
import { store, SheetState } from '../model/store'
import { ShapeDef, SheetLayout } from '../model/types'
import { anchorAt, axesOf, floatsOf } from './arrange'
import { setUi, showToast } from '../chrome/ui'
import { shapesOf } from './shapes'

export type Pt = [number, number]
export interface Pen { name: string; color: string; width: number; alpha?: number }
export type InkTool = { kind: 'select' } | { kind: 'lasso' } | { kind: 'eraser' } | { kind: 'pen'; pen: Pen }

export const DEFAULT_PENS: Pen[] = [
  { name: 'Pen: Black', color: '#000000', width: 2 }, { name: 'Pen: Red', color: '#ff0000', width: 2 },
  { name: 'Pencil', color: '#595959', width: 1.5 }, { name: 'Highlighter: Yellow', color: '#ffff00', width: 14, alpha: 0.5 },
  { name: 'Pen: Blue', color: '#0070c0', width: 2 }, { name: 'Pen: Green', color: '#00b050', width: 2 }
]

const dist = (a: Pt, b: Pt): number => Math.hypot(a[0] - b[0], a[1] - b[1])

/** Ramer-Douglas-Peucker: drops points closer than `eps` to the line through their neighbours. */
export function simplify(pts: Pt[], eps: number): Pt[] {
  if (pts.length < 3) return pts
  const [a, b] = [pts[0], pts[pts.length - 1]]
  const len = dist(a, b)
  let worst = 0, at = 0
  for (let i = 1; i < pts.length - 1; i++) {
    const d = len === 0 ? dist(pts[i], a) : Math.abs((b[0] - a[0]) * (a[1] - pts[i][1]) - (a[0] - pts[i][0]) * (b[1] - a[1])) / len
    if (d > worst) { worst = d; at = i }
  }
  if (worst <= eps) return [a, b]
  return [...simplify(pts.slice(0, at + 1), eps).slice(0, -1), ...simplify(pts.slice(at), eps)]
}

export function bounds(pts: Pt[]): { x: number; y: number; w: number; h: number } {
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1])
  const x = Math.min(...xs), y = Math.min(...ys)
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y }
}

/** The leading `frac` (0..1) of a stroke by point count, with a partial last segment (Ink Replay). */
export function truncate(pts: Pt[], frac: number): Pt[] {
  if (frac >= 1 || pts.length < 2) return pts
  const t = Math.max(0, frac) * (pts.length - 1)
  const i = Math.floor(t), f = t - i
  const out = pts.slice(0, i + 1)
  if (f > 0 && pts[i + 1]) out.push([pts[i][0] + (pts[i + 1][0] - pts[i][0]) * f, pts[i][1] + (pts[i + 1][1] - pts[i][1]) * f])
  return out
}

export const pathD = (pts: Pt[]): string => pts.map((p, i) => `${i ? 'L' : 'M'}${+p[0].toFixed(2)},${+p[1].toFixed(2)}`).join(' ')

/** Ink points in screen px given the ink's current on-screen box. */
export function screenPoints(s: { width: number; height: number; points?: Pt[] }, box: { left: number; top: number; width: number; height: number }): Pt[] {
  const kx = s.width ? box.width / s.width : 1, ky = s.height ? box.height / s.height : 1
  return (s.points ?? []).map((p) => [box.left + p[0] * kx, box.top + p[1] * ky] as Pt)
}

function segDist(p: Pt, a: Pt, b: Pt): number {
  const dx = b[0] - a[0], dy = b[1] - a[1], l2 = dx * dx + dy * dy
  const t = l2 === 0 ? 0 : Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / l2))
  return dist(p, [a[0] + t * dx, a[1] + t * dy])
}

/** True when the eraser circle at `at` touches the polyline. */
export function strokeHit(pts: Pt[], at: Pt, radius: number): boolean {
  if (pts.length === 1) return dist(pts[0], at) <= radius
  for (let i = 0; i < pts.length - 1; i++) if (segDist(at, pts[i], pts[i + 1]) <= radius) return true
  return false
}

export function pointInPolygon(p: Pt, poly: Pt[]): boolean {
  let inside = false
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const [xi, yi] = poly[i], [xj, yj] = poly[j]
    if (yi > p[1] !== yj > p[1] && p[0] < ((xj - xi) * (p[1] - yi)) / (yj - yi) + xi) inside = !inside
  }
  return inside
}

export type Recognized = { kind: 'line'; from: Pt; to: Pt } | { kind: 'rect' | 'ellipse' }

/** Snaps a stroke to a line, rectangle or ellipse, or null when it is none of them (Ink to Shape). */
export function recognizeShape(pts: Pt[]): Recognized | null {
  if (pts.length < 2) return null
  let len = 0
  for (let i = 1; i < pts.length; i++) len += dist(pts[i - 1], pts[i])
  const chord = dist(pts[0], pts[pts.length - 1])
  if (len < 8) return null
  if (chord / len > 0.9) return { kind: 'line', from: pts[0], to: pts[pts.length - 1] }
  if (chord / len > 0.3) return null // an open curve: neither a line nor a closed outline
  const b = bounds(pts)
  if (b.w < 6 || b.h < 6) return null
  let eErr = 0, rErr = 0
  for (const p of pts) {
    const x = ((p[0] - b.x) / b.w) * 2 - 1, y = ((p[1] - b.y) / b.h) * 2 - 1
    eErr += Math.abs(Math.hypot(x, y) - 1)
    rErr += Math.min(1 - Math.abs(x), 1 - Math.abs(y))
  }
  eErr /= pts.length; rErr /= pts.length
  if (Math.min(eErr, rErr) > 0.2) return null
  return { kind: eErr < rErr ? 'ellipse' : 'rect' }
}

/** Builds the ink ShapeDef for a stroke given in absolute sheet px (layout coordinates). */
export function makeInk(l: Partial<SheetLayout>, abs: Pt[], pen: Pen, n: number): ShapeDef | null {
  if (!abs.length) return null
  const pts = simplify(abs, 0.6)
  const b = bounds(pts)
  const pad = pen.width / 2
  const x = Math.max(0, b.x - pad), y = Math.max(0, b.y - pad)
  const width = Math.max(1, b.w + pad * 2), height = Math.max(1, b.h + pad * 2)
  return {
    id: 'shp_ink_' + Date.now().toString(36) + n, name: `Ink ${n}`, kind: 'ink', anchor: anchorAt(x, y, axesOf(l)), width, height,
    points: pts.map((p) => [p[0] - x, p[1] - y] as Pt), line: pen.color, lineWidth: pen.width, ...(pen.alpha !== undefined ? { alpha: pen.alpha } : {})
  }
}

export const inksOf = (sheet: SheetState): ShapeDef[] => shapesOf(sheet).filter((s) => s.kind === 'ink')

export function addInk(sheet: SheetState, abs: Pt[], pen: Pen): void {
  if (!store.allows(sheet, 'objects')) return
  const ink = makeInk(sheet.layout, abs, pen, inksOf(sheet).length + 1)
  if (ink) store.updateLayout(sheet, 'Ink', (l) => { l.shapes = [...(l.shapes ?? []), ink] })
}

export function eraseInks(sheet: SheetState, ids: string[]): void {
  if (!ids.length || !store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, 'Erase Ink', (l) => { l.shapes = (l.shapes ?? []).filter((s) => !ids.includes(s.id)) })
}

/** The anchor of the point (px, py) inside shape s's own box. */
function shift(l: Partial<SheetLayout>, s: ShapeDef, px: number, py: number): ShapeDef['anchor'] {
  const ax = axesOf(l)
  return anchorAt(ax.cols.offset(s.anchor.c) + s.anchor.dx + px, ax.rows.offset(s.anchor.r) + s.anchor.dy + py, ax)
}

/** Converts one ink stroke to the shape it resembles (pure; null when unrecognisable). */
export function inkAsShape(l: Partial<SheetLayout>, s: ShapeDef): ShapeDef | null {
  if (s.kind !== 'ink' || !s.points) return null
  const r = recognizeShape(s.points)
  if (!r) return null
  const { points: _p, alpha: _a, ...rest } = s
  if (r.kind === 'line') {
    const dx = r.to[0] - r.from[0], dy = r.to[1] - r.from[1]
    return { ...rest, kind: 'line', name: s.name.replace('Ink', 'Line'), anchor: shift(l, s, Math.min(r.from[0], r.to[0]), Math.min(r.from[1], r.to[1])), width: Math.max(1, Math.abs(dx)), height: Math.max(1, Math.abs(dy)), ...(dx * dy < 0 ? { flipV: true } : {}) }
  }
  const b = bounds(s.points)
  return { ...rest, kind: r.kind, name: s.name.replace('Ink', r.kind === 'rect' ? 'Rectangle' : 'Oval'), anchor: shift(l, s, b.x, b.y), width: Math.max(8, b.w), height: Math.max(8, b.h) }
}

/** Replaces each selected ink stroke with the line / rectangle / oval it resembles. Returns how many converted. */
export function inkToShape(sheet: SheetState, ids: string[]): number {
  if (!store.allows(sheet, 'objects')) return 0
  let done = 0
  store.updateLayout(sheet, 'Ink to Shape', (l) => {
    l.shapes = (l.shapes ?? []).map((s) => {
      const out = ids.includes(s.id) ? inkAsShape(l, s) : null
      if (out) done++
      return out ?? s
    })
  })
  return done
}

let replayTimer = 0
/** Ink Replay: draws the sheet's strokes one after another, back to front, then shows them all again. */
export function replayInk(sheet: SheetState): void {
  const order = floatsOf(sheet.layout).filter((f) => f.kind === 'shape' && (f.obj as unknown as ShapeDef).kind === 'ink').map((f) => f.id)
  if (!order.length) { showToast('There is no ink on this sheet to replay.'); return }
  window.clearInterval(replayTimer)
  const steps = 12
  let tick = 0
  setUi({ inkReplay: { order, shown: 0, frac: 0 } })
  replayTimer = window.setInterval(() => {
    tick++
    const shown = Math.floor(tick / steps), frac = (tick % steps) / steps
    if (shown >= order.length) { window.clearInterval(replayTimer); setUi({ inkReplay: null }); return }
    setUi({ inkReplay: { order, shown, frac } })
  }, 40)
}

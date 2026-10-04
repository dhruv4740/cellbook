// Page Layout > Arrange on floating objects (charts, pictures, shapes): z-order, align / distribute, group, rotate / flip,
// hide. Pure functions over a layout draft so they run inside store.updateLayout and are unit-testable.
import { ArrangeProps, SheetLayout } from '../model/types'
import { buildAxis, Axis } from '../grid/axis'
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, MAX_COLS, MAX_ROWS } from '../model/types'

export type ObjKind = 'chart' | 'picture' | 'shape'
type Anchored = ArrangeProps & { id: string; anchor: { r: number; c: number; dx: number; dy: number }; width: number; height: number }
export interface FloatRef { id: string; kind: ObjKind; obj: Anchored; name: string }

/** Every floating object on the sheet, back to front. Objects without a z sit above those that have one, in insert order. */
export function floatsOf(l: Partial<SheetLayout>): FloatRef[] {
  const all: FloatRef[] = [
    ...(l.pictures ?? []).map((o) => ({ id: o.id, kind: 'picture' as const, obj: o as Anchored, name: o.name })),
    ...(l.shapes ?? []).map((o) => ({ id: o.id, kind: 'shape' as const, obj: o as Anchored, name: o.name })),
    ...(l.charts ?? []).map((o, i) => ({ id: o.id, kind: 'chart' as const, obj: o as Anchored, name: `Chart ${i + 1}` }))
  ]
  return all.map((f, i) => ({ f, key: f.obj.z ?? 1e9 + i })).sort((a, b) => a.key - b.key).map((x) => x.f)
}

/** Applies a patch per object id across all three lists (ids never collide between kinds). */
function patchAll(l: Partial<SheetLayout>, patches: Map<string, Partial<ArrangeProps> & { anchor?: Anchored['anchor']; width?: number; height?: number }>): void {
  const map = <T extends { id: string }>(list: T[] | undefined): T[] | undefined => list?.map((o) => (patches.has(o.id) ? { ...o, ...patches.get(o.id) } : o))
  if (l.pictures) l.pictures = map(l.pictures)
  if (l.shapes) l.shapes = map(l.shapes)
  if (l.charts) l.charts = map(l.charts)
}

/** The back-to-front id order after moving `ids` (Excel's four Bring / Send commands). */
export function reorderIds(order: string[], ids: string[], mode: 'front' | 'back' | 'forward' | 'backward'): string[] {
  const sel = new Set(ids)
  const out = [...order]
  if (mode === 'front') return [...out.filter((x) => !sel.has(x)), ...out.filter((x) => sel.has(x))]
  if (mode === 'back') return [...out.filter((x) => sel.has(x)), ...out.filter((x) => !sel.has(x))]
  if (mode === 'forward') {
    for (let i = out.length - 2; i >= 0; i--) if (sel.has(out[i]) && !sel.has(out[i + 1])) [out[i], out[i + 1]] = [out[i + 1], out[i]]
  } else {
    for (let i = 1; i < out.length; i++) if (sel.has(out[i]) && !sel.has(out[i - 1])) [out[i], out[i - 1]] = [out[i - 1], out[i]]
  }
  return out
}

export function reorder(l: Partial<SheetLayout>, ids: string[], mode: 'front' | 'back' | 'forward' | 'backward'): void {
  const order = reorderIds(floatsOf(l).map((f) => f.id), ids, mode)
  patchAll(l, new Map(order.map((id, i) => [id, { z: i }])))
}

/** Moves one object to an absolute back-to-front position (Selection Pane arrows). */
export function setOrder(l: Partial<SheetLayout>, order: string[]): void {
  patchAll(l, new Map(order.map((id, i) => [id, { z: i }])))
}

export function axesOf(l: Partial<SheetLayout>): { cols: Axis; rows: Axis } {
  return {
    cols: buildAxis(l.colWidths ?? {}, l.hiddenCols ?? {}, DEFAULT_COL_WIDTH, MAX_COLS),
    rows: buildAxis(l.rowHeights ?? {}, l.hiddenRows ?? {}, DEFAULT_ROW_HEIGHT, MAX_ROWS)
  }
}

export interface AbsRect { x: number; y: number; w: number; h: number }
export const absRect = (o: Anchored, ax: { cols: Axis; rows: Axis }): AbsRect => ({ x: ax.cols.offset(o.anchor.c) + o.anchor.dx, y: ax.rows.offset(o.anchor.r) + o.anchor.dy, w: o.width, h: o.height })
export function anchorAt(x: number, y: number, ax: { cols: Axis; rows: Axis }): Anchored['anchor'] {
  const X = Math.max(0, x), Y = Math.max(0, y)
  const c = ax.cols.indexAt(X), r = ax.rows.indexAt(Y)
  return { r, c, dx: X - ax.cols.offset(c), dy: Y - ax.rows.offset(r) }
}

export type AlignMode = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'

/** New left/top per id for Align (against the selection's bounding box) of two or more objects. */
export function alignRects(rects: Map<string, AbsRect>, mode: AlignMode): Map<string, { x: number; y: number }> {
  const rs = [...rects.values()]
  const minX = Math.min(...rs.map((r) => r.x)), maxX = Math.max(...rs.map((r) => r.x + r.w))
  const minY = Math.min(...rs.map((r) => r.y)), maxY = Math.max(...rs.map((r) => r.y + r.h))
  const out = new Map<string, { x: number; y: number }>()
  for (const [id, r] of rects) {
    let { x, y } = r
    if (mode === 'left') x = minX
    else if (mode === 'right') x = maxX - r.w
    else if (mode === 'center') x = (minX + maxX) / 2 - r.w / 2
    else if (mode === 'top') y = minY
    else if (mode === 'bottom') y = maxY - r.h
    else y = (minY + maxY) / 2 - r.h / 2
    out.set(id, { x, y })
  }
  return out
}

/** Equal gaps between three or more objects; the two outermost stay put. */
export function distributeRects(rects: Map<string, AbsRect>, horizontal: boolean): Map<string, { x: number; y: number }> {
  const items = [...rects.entries()].sort((a, b) => (horizontal ? a[1].x - b[1].x : a[1].y - b[1].y))
  const out = new Map<string, { x: number; y: number }>(items.map(([id, r]) => [id, { x: r.x, y: r.y }]))
  if (items.length < 3) return out
  const size = (r: AbsRect): number => (horizontal ? r.w : r.h)
  const start = horizontal ? items[0][1].x : items[0][1].y
  const last = items[items.length - 1][1]
  const end = (horizontal ? last.x : last.y) + size(last)
  const gap = (end - start - items.reduce((s, [, r]) => s + size(r), 0)) / (items.length - 1)
  let pos = start
  for (const [id, r] of items) {
    out.set(id, horizontal ? { x: pos, y: r.y } : { x: r.x, y: pos })
    pos += size(r) + gap
  }
  return out
}

function moveTo(l: Partial<SheetLayout>, pos: Map<string, { x: number; y: number }>): void {
  const ax = axesOf(l)
  patchAll(l, new Map([...pos].map(([id, p]) => [id, { anchor: anchorAt(p.x, p.y, ax) }])))
}

export function align(l: Partial<SheetLayout>, ids: string[], mode: AlignMode): void {
  const ax = axesOf(l)
  const rects = new Map(floatsOf(l).filter((f) => ids.includes(f.id)).map((f) => [f.id, absRect(f.obj, ax)]))
  if (rects.size >= 2) moveTo(l, alignRects(rects, mode))
}

export function distribute(l: Partial<SheetLayout>, ids: string[], horizontal: boolean): void {
  const ax = axesOf(l)
  const rects = new Map(floatsOf(l).filter((f) => ids.includes(f.id)).map((f) => [f.id, absRect(f.obj, ax)]))
  if (rects.size >= 3) moveTo(l, distributeRects(rects, horizontal))
}

/** Shifts every listed object by (dx, dy) logical px (dragging one member of a group or multi-selection). */
export function shiftBy(l: Partial<SheetLayout>, ids: string[], dx: number, dy: number): void {
  const ax = axesOf(l)
  moveTo(l, new Map(floatsOf(l).filter((f) => ids.includes(f.id)).map((f) => { const r = absRect(f.obj, ax); return [f.id, { x: r.x + dx, y: r.y + dy }] })))
}

/** All ids in the same groups as `ids` (selecting one member selects the whole group). */
export function withGroupMembers(l: Partial<SheetLayout>, ids: string[]): string[] {
  const all = floatsOf(l)
  const groups = new Set(all.filter((f) => ids.includes(f.id) && f.obj.group).map((f) => f.obj.group))
  return all.filter((f) => ids.includes(f.id) || (f.obj.group && groups.has(f.obj.group))).map((f) => f.id)
}

/** Groups the objects (nested groups are flattened into one). Needs two or more. */
export function group(l: Partial<SheetLayout>, ids: string[], groupId: string): void {
  const members = withGroupMembers(l, ids)
  if (members.length >= 2) patchAll(l, new Map(members.map((id) => [id, { group: groupId }])))
}

export function ungroup(l: Partial<SheetLayout>, ids: string[]): void {
  const members = withGroupMembers(l, ids)
  patchAll(l, new Map(members.map((id) => [id, { group: undefined }])))
}

const isLine = (kind: ObjKind, o: unknown): boolean => kind === 'shape' && ['line', 'arrow'].includes((o as { kind: string }).kind)

/** Rotate by 90 degree steps or flip; charts can't rotate or flip (as in Excel). */
export function transform(l: Partial<SheetLayout>, ids: string[], op: { rotate: number } | { flip: 'h' | 'v' }): void {
  const patches = new Map<string, Partial<ArrangeProps>>()
  for (const f of floatsOf(l)) {
    if (!ids.includes(f.id) || f.kind === 'chart') continue
    const o = f.obj
    if ('rotate' in op) patches.set(f.id, { rotation: (((o.rotation ?? 0) + op.rotate) % 360 + 360) % 360 })
    // A diagonal line's horizontal flip looks the same as its vertical flip, so lines only use flipV.
    else if (op.flip === 'v' || isLine(f.kind, o)) patches.set(f.id, { flipV: !o.flipV })
    else patches.set(f.id, { flipH: !o.flipH })
  }
  patchAll(l, patches)
}

export function setHidden(l: Partial<SheetLayout>, ids: string[], hidden: boolean): void {
  patchAll(l, new Map(ids.map((id) => [id, { hidden: hidden || undefined }])))
}

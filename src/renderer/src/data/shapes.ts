// Insert > Shapes / Text Box: floating drawing objects. Every change goes through store.updateLayout so it undoes.
import { store, SheetState } from '../model/store'
import { ShapeDef, ShapeKind } from '../model/types'
import { setUi } from '../chrome/ui'

export const shapesOf = (sheet: SheetState): ShapeDef[] => sheet.layout.shapes ?? []

export const SHAPE_LABELS: Record<ShapeKind, string> = {
  rect: 'Rectangle', roundRect: 'Rectangle: Rounded Corners', ellipse: 'Oval', triangle: 'Isosceles Triangle', diamond: 'Diamond',
  rightArrow: 'Arrow: Right', line: 'Line', arrow: 'Arrow', textbox: 'Text Box'
}

const DEFAULT_SIZE: Record<ShapeKind, [number, number]> = {
  rect: [120, 80], roundRect: [120, 80], ellipse: [120, 80], triangle: [100, 90], diamond: [110, 90], rightArrow: [130, 70], line: [140, 0.01], arrow: [140, 0.01], textbox: [150, 60]
}

/** Excel's default style for an inserted shape: blue fill, darker outline, white text; a text box has no fill or outline. */
export function insertShape(kind: ShapeKind, patch: Partial<ShapeDef> = {}, label?: string): void {
  const sheet = store.active
  if (!store.allows(sheet, 'objects')) return
  const { r, c } = sheet.selection.active
  const [w, h] = DEFAULT_SIZE[kind]
  const n = shapesOf(sheet).filter((s) => (s.kind === 'textbox') === (kind === 'textbox')).length + 1
  const isLine = kind === 'line' || kind === 'arrow'
  const shape: ShapeDef = {
    id: 'shp_' + Date.now().toString(36) + n, name: `${kind === 'textbox' ? 'TextBox' : SHAPE_LABELS[kind].replace(/:.*$/, '')} ${n}`, kind,
    anchor: { r, c, dx: 0, dy: 0 }, width: w, height: Math.max(h, isLine ? 1 : h),
    ...(kind === 'textbox' ? { line: undefined, fontSize: 11 } : isLine ? { line: '#4472c4', lineWidth: 2 } : { fill: '#4472c4', line: '#2f528f', lineWidth: 1, color: '#ffffff', fontSize: 11, align: 'ctr' as const })
  }
  if (isLine) shape.height = 60 // a diagonal that is easy to grab
  Object.assign(shape, patch)
  store.updateLayout(sheet, label ?? (kind === 'textbox' ? 'Insert Text Box' : 'Insert Shape'), (l) => { l.shapes = [...(l.shapes ?? []), shape] })
  setUi({ selectedPicture: shape.id, selectedChart: null })
}

export function updateShape(sheet: SheetState, id: string, patch: Partial<ShapeDef>, label = 'Format Shape'): void {
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, label, (l) => { l.shapes = (l.shapes ?? []).map((p) => (p.id === id ? { ...p, ...patch } : p)) })
}

export function deleteShape(sheet: SheetState, id: string): void {
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, 'Delete Shape', (l) => { l.shapes = (l.shapes ?? []).filter((p) => p.id !== id) })
}

export function reorderShape(sheet: SheetState, id: string, toFront: boolean): void {
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, toFront ? 'Bring to Front' : 'Send to Back', (l) => {
    const list = l.shapes ?? []
    const p = list.find((x) => x.id === id)
    if (p) l.shapes = toFront ? [...list.filter((x) => x.id !== id), p] : [p, ...list.filter((x) => x.id !== id)]
  })
}

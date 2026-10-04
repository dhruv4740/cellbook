import { describe, expect, it } from 'vitest'
import { Pt, bounds, inkAsShape, makeInk, pathD, pointInPolygon, recognizeShape, screenPoints, simplify, strokeHit, truncate } from './ink'
import { injectPictures, parseDrawingObjects } from '../io/xlsxPictures'
import { floatsOf, reorder } from './arrange'

const circle = (n = 40): Pt[] => Array.from({ length: n + 1 }, (_, i) => [50 + 40 * Math.cos((i / n) * 2 * Math.PI), 50 + 40 * Math.sin((i / n) * 2 * Math.PI)] as Pt)
const square = (): Pt[] => {
  const out: Pt[] = []
  const corners: Pt[] = [[0, 0], [80, 0], [80, 60], [0, 60], [0, 0]]
  for (let k = 0; k < 4; k++) for (let t = 0; t < 10; t++) out.push([corners[k][0] + ((corners[k + 1][0] - corners[k][0]) * t) / 10, corners[k][1] + ((corners[k + 1][1] - corners[k][1]) * t) / 10])
  out.push([0, 0])
  return out
}

describe('ink geometry', () => {
  it('simplify drops collinear points but keeps corners', () => {
    expect(simplify([[0, 0], [1, 0.01], [2, 0], [3, 0]], 0.5)).toEqual([[0, 0], [3, 0]])
    expect(simplify([[0, 0], [5, 0], [5, 5]], 0.5)).toHaveLength(3)
  })
  it('truncate returns a leading fraction with an interpolated end', () => {
    expect(truncate([[0, 0], [10, 0]], 0.5)).toEqual([[0, 0], [5, 0]])
    expect(truncate([[0, 0], [10, 0]], 1)).toEqual([[0, 0], [10, 0]])
  })
  it('hit-tests the eraser against segments', () => {
    expect(strokeHit([[0, 0], [100, 0]], [50, 4], 5)).toBe(true)
    expect(strokeHit([[0, 0], [100, 0]], [50, 9], 5)).toBe(false)
  })
  it('point in polygon (lasso)', () => {
    const tri: Pt[] = [[0, 0], [10, 0], [0, 10]]
    expect(pointInPolygon([2, 2], tri)).toBe(true)
    expect(pointInPolygon([9, 9], tri)).toBe(false)
  })
  it('builds svg path data and scales points to a box', () => {
    expect(pathD([[0, 0], [1.234, 2]])).toBe('M0,0 L1.23,2')
    expect(screenPoints({ width: 10, height: 10, points: [[5, 5]] }, { left: 100, top: 0, width: 20, height: 20 })).toEqual([[110, 10]])
  })
})

describe('Ink to Shape', () => {
  it('recognises a line, a rectangle and an oval', () => {
    expect(recognizeShape([[0, 0], [30, 31], [60, 60]])?.kind).toBe('line')
    expect(recognizeShape(square())?.kind).toBe('rect')
    expect(recognizeShape(circle())?.kind).toBe('ellipse')
  })
  it('rejects scribbles and open curves', () => {
    expect(recognizeShape([[0, 0], [1, 1]])).toBeNull()
    const arc = circle().slice(0, 25)
    expect(recognizeShape(arc)).toBeNull()
  })
  it('turns an ink object into a plain shape, keeping id and colour', () => {
    const ink = makeInk({}, circle().map((p) => [p[0] + 100, p[1] + 40] as Pt), { name: 'p', color: '#ff0000', width: 2 }, 1)!
    const s = inkAsShape({}, ink)!
    expect(s.kind).toBe('ellipse')
    expect(s.id).toBe(ink.id)
    expect(s.line).toBe('#ff0000')
    expect(s.points).toBeUndefined()
    expect(s.width).toBeGreaterThan(70)
  })
  it('flips a rising line', () => {
    const ink = makeInk({}, [[10, 100], [60, 50], [110, 0]], { name: 'p', color: '#000000', width: 2 }, 1)!
    expect(inkAsShape({}, ink)).toMatchObject({ kind: 'line', flipV: true })
  })
})

describe('ink objects', () => {
  it('stores points relative to the anchored box and joins arrange z-order', () => {
    const ink = makeInk({}, [[100, 50], [140, 90]], { name: 'p', color: '#000000', width: 4, alpha: 0.5 }, 2)!
    expect(ink.kind).toBe('ink')
    expect(ink.alpha).toBe(0.5)
    const b = bounds(ink.points!)
    expect(b.x).toBeGreaterThanOrEqual(0)
    const l = { shapes: [ink, { ...ink, id: 'shp_b', name: 'Ink 3' }] }
    reorder(l as never, ['shp_b'], 'back')
    expect(floatsOf(l as never).map((f) => f.id)).toEqual(['shp_b', ink.id])
  })
  it('survives the .xlsx drawing round trip', () => {
    const ink = makeInk({}, [[30, 20], [60, 70], [100, 30]], { name: 'p', color: '#0070c0', width: 3, alpha: 0.5 }, 1)!
    const enc = new TextEncoder()
    const out = injectPictures({ files: { '[Content_Types].xml': enc.encode('<Types></Types>'), 'xl/worksheets/sheet1.xml': enc.encode('<worksheet><sheetData/></worksheet>') } } as never, [[]] as never, [{ shapes: [ink] }] as never)
    const files = (out as unknown as { files: Record<string, Uint8Array> }).files
    const dec = new TextDecoder()
    const back = parseDrawingObjects(files, dec, 'xl/worksheets/sheet1.xml', dec.decode(files['xl/worksheets/sheet1.xml'])).shapes
    expect(back).toHaveLength(1)
    expect(back[0].kind).toBe('ink')
    expect(back[0].line).toBe('#0070c0')
    expect(back[0].alpha).toBeCloseTo(0.5)
    expect(back[0].lineWidth).toBeCloseTo(3)
    expect(back[0].points).toHaveLength(3)
    expect(back[0].points![1][0]).toBeCloseTo(ink.points![1][0], 0)
    expect(back[0].points![1][1]).toBeCloseTo(ink.points![1][1], 0)
  })
})

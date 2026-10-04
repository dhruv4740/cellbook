import { describe, it, expect } from 'vitest'
import { SheetLayout } from '../model/types'
import { floatsOf, reorderIds, reorder, align, distribute, group, ungroup, transform, setHidden, withGroupMembers, shiftBy, axesOf, absRect } from './arrange'

const box = (id: string, x: number, y: number, w = 40, h = 20) => ({ id, name: id, anchor: { r: 0, c: 0, dx: x, dy: y }, width: w, height: h })
const layout = (): Partial<SheetLayout> => ({
  pictures: [box('p1', 0, 0)] as never, shapes: [{ ...box('s1', 100, 10), kind: 'rect' }, { ...box('s2', 300, 30, 40, 40), kind: 'line' }] as never,
  charts: [box('c1', 50, 50)] as never
})
const order = (l: Partial<SheetLayout>): string[] => floatsOf(l).map((f) => f.id)

describe('z-order', () => {
  it('defaults to pictures, shapes, then charts', () => expect(order(layout())).toEqual(['p1', 's1', 's2', 'c1']))
  it('reorders per Excel (forward passes one non-selected, front/back go to the ends)', () => {
    const o = ['a', 'b', 'c', 'd']
    expect(reorderIds(o, ['a'], 'forward')).toEqual(['b', 'a', 'c', 'd'])
    expect(reorderIds(o, ['d'], 'forward')).toEqual(o)
    expect(reorderIds(o, ['c'], 'backward')).toEqual(['a', 'c', 'b', 'd'])
    expect(reorderIds(o, ['a', 'b'], 'forward')).toEqual(['c', 'a', 'b', 'd'])
    expect(reorderIds(o, ['b'], 'front')).toEqual(['a', 'c', 'd', 'b'])
    expect(reorderIds(o, ['c', 'd'], 'back')).toEqual(['c', 'd', 'a', 'b'])
  })
  it('writes z across all three kinds', () => {
    const l = layout()
    reorder(l, ['c1'], 'back')
    expect(order(l)).toEqual(['c1', 'p1', 's1', 's2'])
    reorder(l, ['p1'], 'front')
    expect(order(l)).toEqual(['c1', 's1', 's2', 'p1'])
  })
})

describe('align / distribute', () => {
  it('aligns to the selection edge, centre or middle', () => {
    const l = layout()
    align(l, ['p1', 's1'], 'right')
    const ax = axesOf(l)
    expect(absRect(l.pictures![0] as never, ax).x).toBe(100) // s1 right edge 140, p1 width 40
    align(l, ['p1', 's1'], 'top')
    expect(absRect(l.shapes![0] as never, ax).y).toBe(0)
    align(l, ['p1', 's1'], 'middle')
    expect(absRect(l.pictures![0] as never, ax).y).toBe(absRect(l.shapes![0] as never, ax).y)
  })
  it('distributes three objects with equal gaps and leaves the outer two in place', () => {
    const l = layout()
    distribute(l, ['p1', 's1', 's2'], true) // x 0..40, 100..140, 300..340 -> gaps of 110
    const ax = axesOf(l)
    expect(l.shapes!.map((s) => absRect(s as never, ax).x)).toEqual([150, 300])
    expect(absRect(l.pictures![0] as never, ax).x).toBe(0)
  })
  it('does nothing with fewer objects than needed', () => {
    const l = layout()
    distribute(l, ['p1', 's1'], true)
    expect(absRect(l.shapes![0] as never, axesOf(l)).x).toBe(100)
  })
  it('shifts a selection by the same delta', () => {
    const l = layout()
    shiftBy(l, ['p1', 's1'], 10, 5)
    const ax = axesOf(l)
    expect(absRect(l.pictures![0] as never, ax)).toMatchObject({ x: 10, y: 5 })
    expect(absRect(l.shapes![0] as never, ax)).toMatchObject({ x: 110, y: 15 })
  })
})

describe('group / rotate / hide', () => {
  it('groups, expands selection to the whole group and ungroups', () => {
    const l = layout()
    group(l, ['p1', 's1'], 'g1')
    expect(withGroupMembers(l, ['s1'])).toEqual(['p1', 's1'])
    group(l, ['s1', 'c1'], 'g2') // merges into one flat group
    expect(withGroupMembers(l, ['c1'])).toEqual(['p1', 's1', 'c1'])
    ungroup(l, ['p1'])
    expect(withGroupMembers(l, ['p1'])).toEqual(['p1'])
  })
  it('rotates in 90 degree steps and flips; charts are skipped; line flips use flipV', () => {
    const l = layout()
    transform(l, ['p1', 'c1'], { rotate: 90 })
    transform(l, ['p1'], { rotate: -180 })
    expect((l.pictures![0] as { rotation?: number }).rotation).toBe(270)
    expect((l.charts![0] as { rotation?: number }).rotation).toBeUndefined()
    transform(l, ['p1', 's1', 's2'], { flip: 'h' })
    expect(l.pictures![0]).toMatchObject({ flipH: true })
    expect(l.shapes![0]).toMatchObject({ flipH: true })
    expect(l.shapes![1]).toMatchObject({ flipV: true }) // line
  })
  it('hides and shows', () => {
    const l = layout()
    setHidden(l, ['s1'], true)
    expect(floatsOf(l).find((f) => f.id === 's1')!.obj.hidden).toBe(true)
    setHidden(l, ['s1'], false)
    expect(floatsOf(l).find((f) => f.id === 's1')!.obj.hidden).toBeUndefined()
  })
})

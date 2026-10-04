import { describe, expect, it } from 'vitest'
import { addView, deleteView, listViews, showView } from './customViews'
import { store } from '../model/store'

describe('custom views', () => {
  it('captures and restores zoom, freeze, flags and selection, replacing same-name views', () => {
    const s = store.active
    store.updateLayout(s, 'x', (l) => { l.zoom = 150; l.frozen = { rows: 2, cols: 1 }; l.showGridlines = false })
    store.setSelection({ ranges: [{ r1: 3, c1: 2, r2: 4, c2: 3 }], active: { r: 3, c: 2 } }, s)
    addView('A')
    store.updateLayout(s, 'x', (l) => { l.zoom = 100; l.frozen = { rows: 0, cols: 0 }; l.showGridlines = true })
    store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }, s)
    addView('a')
    expect(listViews()).toHaveLength(1)
    addView('B')
    store.updateLayout(s, 'x', (l) => { l.zoom = 150 })
    expect(showView(listViews().find((v) => v.name === 'B')!)).toBe(true)
    expect(store.active.layout.zoom).toBe(100)
    expect(store.active.layout.showGridlines).toBe(true)
    deleteView('B'); deleteView('a')
    expect(listViews()).toHaveLength(0)
  })
})

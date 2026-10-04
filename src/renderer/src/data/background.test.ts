import { describe, it, expect } from 'vitest'
import { store } from '../model/store'
import { setSheetBackground, backgroundVars } from './background'

describe('sheet background', () => {
  it('sets, deletes and undoes as layout changes', () => {
    const sh = store.active
    setSheetBackground(sh, 'data:image/png;base64,AAAA')
    expect(sh.layout.background).toBe('data:image/png;base64,AAAA')
    setSheetBackground(sh, null)
    expect(sh.layout.background).toBeUndefined()
    store.undo()
    expect(sh.layout.background).toBe('data:image/png;base64,AAAA')
    store.undo()
    expect(sh.layout.background).toBeUndefined()
  })
  it('exposes CSS vars only when a background exists', () => {
    expect(backgroundVars(undefined, 40, 20, { x: 0, y: 0 })).toEqual({})
    expect(backgroundVars('data:x', 40, 20, { x: 10, y: 5 })).toMatchObject({ '--xl-bg-x': '40px', '--xl-bg-sx': '-10px', '--xl-bg-sy': '-5px' })
  })
})

import { describe, it, expect } from 'vitest'
import { canHideWindow, hiddenWindows } from './windowVisibility'

const w = (id: number, visible: boolean) => ({ id, visible, title: 'Book' + id })

describe('window hide / unhide', () => {
  it('refuses to hide the only visible window', () => {
    expect(canHideWindow([w(1, true)], 1)).toBe(false)
    expect(canHideWindow([w(1, true), w(2, false)], 1)).toBe(false)
  })
  it('allows hiding when another window stays visible', () => {
    expect(canHideWindow([w(1, true), w(2, true)], 1)).toBe(true)
  })
  it('lists only hidden windows', () => {
    expect(hiddenWindows([w(1, true), w(2, false)])).toEqual([{ id: 2, title: 'Book2' }])
  })
})

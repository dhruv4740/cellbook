import { describe, it, expect, beforeEach } from 'vitest'
import { store } from '../model/store'
import { loadSpell, findMisspelling, suggestions, ignoreAll, resetIgnored } from './spell'

describe('spelling', () => {
  beforeEach(() => { store.newWorkbook(); resetIgnored() })
  it('finds misspelled words in text cells, skipping caps, numbers, urls and formulas', async () => {
    const sp = await loadSpell()
    const sh = store.active
    store.setRaw(sh, 0, 0, 'The quick brown fox')
    store.setRaw(sh, 1, 0, 'Shiping adress: NYC 12ab https://exmple.com')
    store.setRaw(sh, 2, 0, '=A1&"wrongg"')
    store.setRaw(sh, 3, 0, 'teh end')
    const b = { r1: 0, c1: 0, r2: 5, c2: 0 }
    const first = findMisspelling(sp, sh, b, { r: 0, c: 0, index: 0 })!
    expect([first.r, first.word]).toEqual([1, 'Shiping'])
    const second = findMisspelling(sp, sh, b, { r: 1, c: 0, index: first.index + first.word.length })!
    expect(second.word).toBe('adress')
    const third = findMisspelling(sp, sh, b, { r: 1, c: 0, index: second.index + second.word.length })!
    expect([third.r, third.word]).toEqual([3, 'teh'])
    expect(suggestions(sp, 'teh')).toContain('the')
    ignoreAll('teh')
    expect(findMisspelling(sp, sh, b, { r: 1, c: 0, index: second.index + second.word.length })).toBeNull()
  })
})

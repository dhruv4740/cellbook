import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { removeDuplicates } from './removeDuplicates'

describe('removeDuplicates', () => {
  it('identifies and removes duplicate rows based on specified columns', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 0, raw: 'First' },
      { r: 0, c: 1, raw: 'Last' },
      { r: 0, c: 2, raw: 'Score' },
      { r: 1, c: 0, raw: 'John' },
      { r: 1, c: 1, raw: 'Doe' },
      { r: 1, c: 2, raw: '10' },
      { r: 2, c: 0, raw: 'Jane' },
      { r: 2, c: 1, raw: 'Smith' },
      { r: 2, c: 2, raw: '20' },
      { r: 3, c: 0, raw: 'John' },
      { r: 3, c: 1, raw: 'Doe' },
      { r: 3, c: 2, raw: '30' }
    ], 'Test')

    // Checking only columns 0 and 1 (First & Last)
    const res = removeDuplicates(store, sheet, { r1: 0, c1: 0, r2: 3, c2: 2 }, [0, 1], true)
    expect(res.removedCount).toBe(1)
    expect(res.remainingCount).toBe(2)

    // Row 1: John, Doe, 10
    // Row 2: Jane, Smith, 20
    // Row 3: cleared
    expect(store.getRaw(sheet, 1, 0)).toBe('John')
    expect(store.getRaw(sheet, 2, 0)).toBe('Jane')
    expect(store.getRaw(sheet, 3, 0)).toBe('')
  })
})

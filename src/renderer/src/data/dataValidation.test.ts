import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { validateInput, addDataValidation, getListItems } from './dataValidation'

describe('dataValidation', () => {
  it('validates whole number between 10 and 100', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    addDataValidation(store, sheet, {
      ranges: [{ r1: 0, c1: 0, r2: 5, c2: 0 }],
      allow: 'whole',
      operator: 'between',
      formula1: '10',
      formula2: '100'
    })

    expect(validateInput(store, sheet, 0, 0, '50').ok).toBe(true)
    expect(validateInput(store, sheet, 0, 0, '5').ok).toBe(false)
    expect(validateInput(store, sheet, 0, 0, '150').ok).toBe(false)
    expect(validateInput(store, sheet, 0, 0, 'abc').ok).toBe(false)
  })

  it('validates list options', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    addDataValidation(store, sheet, {
      ranges: [{ r1: 0, c1: 0, r2: 5, c2: 0 }],
      allow: 'list',
      formula1: 'Red, Green, Blue'
    })

    expect(validateInput(store, sheet, 0, 0, 'Red').ok).toBe(true)
    expect(validateInput(store, sheet, 0, 0, 'green').ok).toBe(true)
    expect(validateInput(store, sheet, 0, 0, 'Yellow').ok).toBe(false)
  })

  it('extracts list items from cell range', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setCells(sheet, [
      { r: 0, c: 1, raw: 'North' },
      { r: 1, c: 1, raw: 'South' },
      { r: 2, c: 1, raw: 'East' },
      { r: 3, c: 1, raw: 'West' }
    ], 'Test')

    const rule = {
      id: 'dv1',
      ranges: [{ r1: 0, c1: 0, r2: 5, c2: 0 }],
      allow: 'list' as const,
      formula1: '=B1:B4'
    }

    const items = getListItems(store, sheet, rule)
    expect(items).toEqual(['North', 'South', 'East', 'West'])
  })
})

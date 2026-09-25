import { describe, it, expect } from 'vitest'
import { WorkbookStore } from '../model/store'
import { evaluateCf, addCfRule, clearCfRules, interpolateColor } from './conditional'

describe('conditional formatting evaluation', () => {
  it('evaluates cellIs greaterThan rule and applies style', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setRaw(sheet, 0, 0, '15')
    store.setRaw(sheet, 1, 0, '5')

    addCfRule(store, sheet, {
      type: 'cellIs',
      operator: 'greaterThan',
      value: '10',
      ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }],
      format: { fill: '#FFC7CE', color: '#9C0006' }
    })

    const cfA1 = evaluateCf(store, sheet, 0, 0)
    expect(cfA1.style?.fill).toBe('#FFC7CE')
    expect(cfA1.style?.color).toBe('#9C0006')

    const cfA2 = evaluateCf(store, sheet, 1, 0)
    expect(cfA2.style).toBeUndefined()
  })

  it('evaluates textContains and duplicateValues', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setRaw(sheet, 0, 0, 'Completed')
    store.setRaw(sheet, 1, 0, 'Pending')
    store.setRaw(sheet, 2, 0, 'Completed')

    // textContains
    addCfRule(store, sheet, {
      type: 'textContains',
      text: 'Pending',
      ranges: [{ r1: 0, c1: 0, r2: 2, c2: 0 }],
      format: { fill: '#FFFFCC' }
    })

    // duplicateValues
    addCfRule(store, sheet, {
      type: 'duplicateValues',
      ranges: [{ r1: 0, c1: 0, r2: 2, c2: 0 }],
      format: { bold: true }
    })

    expect(evaluateCf(store, sheet, 1, 0).style?.fill).toBe('#FFFFCC')
    expect(evaluateCf(store, sheet, 0, 0).style?.bold).toBe(true)
    expect(evaluateCf(store, sheet, 2, 0).style?.bold).toBe(true)
    expect(evaluateCf(store, sheet, 1, 0).style?.bold).toBeUndefined()
  })

  it('evaluates dataBar and colorScale', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setRaw(sheet, 0, 0, '10')
    store.setRaw(sheet, 1, 0, '20')
    store.setRaw(sheet, 2, 0, '30')

    addCfRule(store, sheet, {
      type: 'dataBar',
      dataBar: { color: '#638EC6' },
      ranges: [{ r1: 0, c1: 0, r2: 2, c2: 0 }]
    })

    const evalMin = evaluateCf(store, sheet, 0, 0)
    expect(evalMin.dataBar?.fraction).toBe(0)
    expect(evalMin.dataBar?.color).toBe('#638EC6')

    const evalMax = evaluateCf(store, sheet, 2, 0)
    expect(evalMax.dataBar?.fraction).toBe(1)
  })

  it('evaluates formula expressions relative to cell position', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setRaw(sheet, 0, 0, '10')
    store.setRaw(sheet, 0, 1, '100')
    store.setRaw(sheet, 1, 0, '20')
    store.setRaw(sheet, 1, 1, '50')

    // Rule: B1 > A1
    addCfRule(store, sheet, {
      type: 'expression',
      formula: '=B1>A1',
      ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }],
      format: { fill: '#C6EFCE' }
    })

    expect(evaluateCf(store, sheet, 0, 0).style?.fill).toBe('#C6EFCE')
    expect(evaluateCf(store, sheet, 1, 0).style?.fill).toBe('#C6EFCE')
  })

  it('respects stopIfTrue priority', () => {
    const store = new WorkbookStore()
    const sheet = store.active
    store.setRaw(sheet, 0, 0, '100')

    addCfRule(store, sheet, {
      type: 'cellIs',
      operator: 'greaterThan',
      value: '50',
      ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }],
      format: { fill: '#111111' },
      stopIfTrue: true
    })

    addCfRule(store, sheet, {
      type: 'cellIs',
      operator: 'greaterThan',
      value: '10',
      ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }],
      format: { fill: '#222222' }
    })

    expect(evaluateCf(store, sheet, 0, 0).style?.fill).toBe('#111111')
  })
})

import { describe, expect, it } from 'vitest'
import { defaultLayout } from './defaults'
import { WorkbookStore } from './store'
import { groupLines, outlineRuns, setRunCollapsed, showLevels, ungroupLines, MAX_OUTLINE_LEVEL } from './outline'
import { readXlsx, writeXlsx } from '../io/xlsx'

describe('outline model', () => {
  it('builds nested runs with summary lines and collapsed state', () => {
    const l = defaultLayout()
    groupLines(l, 'row', 1, 3); groupLines(l, 'row', 2, 2)
    const runs = outlineRuns(l, 'row')
    expect(runs).toEqual([
      { level: 1, start: 1, end: 3, summary: 4, collapsed: false },
      { level: 2, start: 2, end: 2, summary: 3, collapsed: false }
    ])
    setRunCollapsed(l, 'row', runs[0], true)
    expect(Object.keys(l.hiddenRows).map(Number)).toEqual([1, 2, 3])
    expect(outlineRuns(l, 'row')[0].collapsed).toBe(true)
  })
  it('expanding an outer group keeps a collapsed inner group closed', () => {
    const l = defaultLayout()
    groupLines(l, 'row', 1, 4); groupLines(l, 'row', 2, 3)
    const inner = outlineRuns(l, 'row')[1]
    setRunCollapsed(l, 'row', inner, true)
    setRunCollapsed(l, 'row', outlineRuns(l, 'row')[0], true)
    setRunCollapsed(l, 'row', outlineRuns(l, 'row')[0], false)
    expect(Object.keys(l.hiddenRows).map(Number)).toEqual([2, 3])
  })
  it('show levels buttons and ungroup', () => {
    const l = defaultLayout()
    groupLines(l, 'col', 1, 3); groupLines(l, 'col', 2, 2)
    showLevels(l, 'col', 1)
    expect(Object.keys(l.hiddenCols).map(Number)).toEqual([1, 2, 3])
    showLevels(l, 'col', 2)
    expect(Object.keys(l.hiddenCols).map(Number)).toEqual([2])
    showLevels(l, 'col', 3)
    expect(l.hiddenCols).toEqual({})
    ungroupLines(l, 'col', 1, 3)
    expect(l.colLevels).toEqual({ 2: 1 })
  })
  it('stops at the maximum depth', () => {
    const l = defaultLayout()
    for (let i = 0; i < MAX_OUTLINE_LEVEL; i++) expect(groupLines(l, 'row', 0, 0)).toBe(true)
    expect(groupLines(l, 'row', 0, 0)).toBe(false)
  })
})

describe('outline in the store and files', () => {
  it('undo, row insert inside a group, and save/reload', async () => {
    const s = new WorkbookStore()
    const sh = s.active
    for (let r = 0; r < 8; r++) s.setRaw(sh, r, 0, String(r))
    s.updateLayout(sh, 'Group', (l) => { groupLines(l, 'row', 1, 3); groupLines(l, 'col', 1, 2) })
    s.updateLayout(sh, 'Hide Detail', (l) => setRunCollapsed(l, 'row', outlineRuns(l, 'row')[0], true))
    s.undo()
    expect(sh.layout.hiddenRows).toEqual({})
    s.redo()
    s.insertRows(sh, 6, 1) // below the group: unaffected
    expect(sh.layout.rowLevels).toEqual({ 1: 1, 2: 1, 3: 1 })
    s.updateLayout(sh, 'Show Detail', (l) => setRunCollapsed(l, 'row', outlineRuns(l, 'row')[0], false))
    s.insertRows(sh, 2, 1) // inside the group: joins it
    expect(sh.layout.rowLevels).toEqual({ 1: 1, 2: 1, 3: 1, 4: 1 })
    s.updateLayout(sh, 'Hide Detail', (l) => setRunCollapsed(l, 'row', outlineRuns(l, 'row')[0], true))
    s.updateLayout(sh, 'Outline', (l) => { l.outline = { summaryBelow: false } })
    const back = await readXlsx(await writeXlsx(s.toData()))
    const l = back.sheets[0].layout!
    expect(l.rowLevels).toEqual({ 1: 1, 2: 1, 3: 1, 4: 1 })
    expect(l.colLevels).toEqual({ 1: 1, 2: 1 })
    expect(Object.keys(l.hiddenRows!).map(Number)).toEqual([1, 2, 3, 4])
    expect(l.outline).toEqual({ summaryBelow: false })
    expect(Object.keys(l.collapsedRows ?? {})).toEqual(['0']) // summary is above the group
  })
})

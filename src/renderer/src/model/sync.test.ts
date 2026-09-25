import { describe, it, expect, beforeEach } from 'vitest'
import { WorkbookStore } from './store'
import { SyncMessage } from './sync'

describe('WorkbookStore multi-window synchronization', () => {
  let storeA: WorkbookStore
  let storeB: WorkbookStore

  beforeEach(() => {
    storeA = new WorkbookStore()
    storeB = new WorkbookStore()

    // Load identical snapshot into both stores (as happens when Window 2 is created)
    const initial = storeA.toData()
    storeB.load(initial, storeA.fileName, storeA.filePath)

    // Connect bidirectional sync channels via IPC mock
    storeA.onBroadcast = (msg: SyncMessage) => {
      storeB.applyRemoteSync(msg)
    }
    storeB.onBroadcast = (msg: SyncMessage) => {
      storeA.applyRemoteSync(msg)
    }
  })

  it('syncs cell edits from storeA to storeB immediately', () => {
    storeA.setRangeRaw(storeA.active, 0, 0, [['Hello World']])
    expect(storeA.getRaw(storeA.active, 0, 0)).toBe('Hello World')
    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('Hello World')
    expect(storeA.canUndo).toBe(true)
    expect(storeB.canUndo).toBe(true)
  })

  it('syncs undo triggered from either window', () => {
    storeA.setRangeRaw(storeA.active, 0, 0, [['Initial']])
    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('Initial')

    // Undo from storeB
    storeB.undo()
    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('')
    expect(storeA.getRaw(storeA.active, 0, 0)).toBe('')

    // Redo from storeA
    storeA.redo()
    expect(storeA.getRaw(storeA.active, 0, 0)).toBe('Initial')
    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('Initial')
  })

  it('syncs styles applied in one window to the other', () => {
    storeA.setRangeRaw(storeA.active, 1, 1, [['100']])
    storeA.applyStyle(storeA.active, [{ r1: 1, c1: 1, r2: 1, c2: 1 }], { bold: true, fill: '#ffff00' })

    const styleB = storeB.getStyle(storeB.active, 1, 1)
    expect(styleB?.bold).toBe(true)
    expect(styleB?.fill).toBe('#ffff00')
  })

  it('syncs compound operations inside transact as a single undo step', () => {
    storeA.transact('Paste', () => {
      storeA.setCells(storeA.active, [
        { r: 0, c: 0, raw: 'A' },
        { r: 0, c: 1, raw: 'B' }
      ], 'Paste')
      storeA.applyStyle(storeA.active, [{ r1: 0, c1: 0, r2: 0, c2: 1 }], { italic: true })
    })

    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('A')
    expect(storeB.getRaw(storeB.active, 0, 1)).toBe('B')
    expect(storeB.getStyle(storeB.active, 0, 0)?.italic).toBe(true)
    expect(storeB.getStyle(storeB.active, 0, 1)?.italic).toBe(true)

    // Should have exactly 1 undo step
    expect(storeB.undoLabels[0]).toBe('Paste')
    storeB.undo()

    expect(storeA.getRaw(storeA.active, 0, 0)).toBe('')
    expect(storeA.getRaw(storeA.active, 0, 1)).toBe('')
    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('')
    expect(storeB.getRaw(storeB.active, 0, 1)).toBe('')
  })

  it('keeps activeSheet independent between windows', () => {
    storeA.addSheet('Sheet2')
    expect(storeA.sheets.length).toBe(2)
    expect(storeB.sheets.length).toBe(2)

    // storeA switched to Sheet2, storeB should still be on Sheet1 (index 0)
    expect(storeA.activeIndex).toBe(1)
    expect(storeB.activeIndex).toBe(0)

    // Editing Sheet2 in storeA syncs to Sheet2 in storeB without moving storeB's activeIndex
    const sheet2A = storeA.sheets[1]
    const sheet2B = storeB.sheets[1]
    storeA.setRangeRaw(sheet2A, 0, 0, [['Sheet2 Data']])
    expect(storeB.getRaw(sheet2B, 0, 0)).toBe('Sheet2 Data')
    expect(storeB.activeIndex).toBe(0)
  })

  it('keeps zoom and freeze panes views independent between windows', () => {
    // Window A zooms to 150%
    storeA.updateLayout(storeA.active, 'Zoom', (l) => { l.zoom = 150 })
    expect(storeA.active.layout.zoom).toBe(150)
    // Window B retains its own zoom (100)
    expect(storeB.active.layout.zoom).toBe(100)

    // Window A freezes panes at row 2, col 1
    storeA.updateLayout(storeA.active, 'Freeze Panes', (l) => { l.frozen = { rows: 2, cols: 1 } })
    expect(storeA.active.layout.frozen).toEqual({ rows: 2, cols: 1 })
    // Window B retains its own freeze panes view
    expect(storeB.active.layout.frozen).toEqual({ rows: 0, cols: 0 })

    // Other layout changes (e.g. column widths) DO sync across windows
    storeA.setColWidth(storeA.active, [3], 180)
    expect(storeB.active.layout.colWidths[3]).toBe(180)
    // And Window B's independent zoom and freeze panes remain untouched
    expect(storeB.active.layout.zoom).toBe(100)
    expect(storeB.active.layout.frozen).toEqual({ rows: 0, cols: 0 })
  })

  it('syncs row and column insertions and deletions', () => {
    storeA.setRangeRaw(storeA.active, 0, 0, [['Row0']])
    storeA.setRangeRaw(storeA.active, 1, 0, [['Row1']])

    // Insert 2 rows at row 1
    storeA.insertRows(storeA.active, 1, 2)
    expect(storeA.getRaw(storeA.active, 0, 0)).toBe('Row0')
    expect(storeA.getRaw(storeA.active, 1, 0)).toBe('')
    expect(storeA.getRaw(storeA.active, 3, 0)).toBe('Row1')

    expect(storeB.getRaw(storeB.active, 0, 0)).toBe('Row0')
    expect(storeB.getRaw(storeB.active, 1, 0)).toBe('')
    expect(storeB.getRaw(storeB.active, 3, 0)).toBe('Row1')

    // Delete the 2 inserted rows
    storeB.deleteRows(storeB.active, 1, 2)
    expect(storeA.getRaw(storeA.active, 1, 0)).toBe('Row1')
    expect(storeB.getRaw(storeB.active, 1, 0)).toBe('Row1')
  })

  it('syncs sheet renaming and moving', () => {
    storeA.addSheet('Expenses')
    expect(storeB.sheets.some((s) => s.name === 'Expenses')).toBe(true)

    // Rename sheet
    const expIdx = storeA.sheets.findIndex((s) => s.name === 'Expenses')
    storeA.renameSheet(expIdx, 'Budget2026')
    expect(storeB.sheets.some((s) => s.name === 'Budget2026')).toBe(true)
    expect(storeB.sheets.some((s) => s.name === 'Expenses')).toBe(false)

    // Move sheet
    storeB.moveSheet(1, 0)
    expect(storeA.sheets[0].name).toBe('Budget2026')
    expect(storeB.sheets[0].name).toBe('Budget2026')
  })
})

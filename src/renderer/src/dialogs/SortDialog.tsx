import { runMaybeBusy } from '../chrome/ui'
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { colName, normRange } from '../model/types'
import { SortLevel, detectHeaders, sortRange } from '../data/sort'
import './SortDialog.css'

export interface SortDialogProps {
  store: WorkbookStore
  onClose: () => void
}

export function SortDialog({ store, onClose }: SortDialogProps) {
  const sheet = store.active
  const range = normRange(sheet.selection.ranges[0] || { r1: 0, c1: 0, r2: 0, c2: 0 })

  const [hasHeaders, setHasHeaders] = useState<boolean>(() => detectHeaders(store, sheet, range))
  const [caseSensitive, setCaseSensitive] = useState(false)
  const [showOptions, setShowOptions] = useState(false)

  // Compute column choices
  const columns: { index: number; name: string }[] = []
  for (let c = range.c1; c <= range.c2; c++) {
    let headerName = ''
    if (hasHeaders) {
      headerName = store.getRaw(sheet, range.r1, c).trim()
    }
    const fallbackName = 'Column ' + colName(c)
    columns.push({ index: c, name: headerName || fallbackName })
  }

  const [levels, setLevels] = useState<SortLevel[]>([
    { columnIndex: range.c1, sortOn: 'values', order: 'asc' }
  ])
  const [selectedIndex, setSelectedIndex] = useState(0)

  const handleAddLevel = () => {
    // Pick first column not yet in levels if possible
    const unused = columns.find((c) => !levels.some((l) => l.columnIndex === c.index))
    const colIdx = unused ? unused.index : range.c1
    const next = [...levels, { columnIndex: colIdx, sortOn: 'values' as const, order: 'asc' as const }]
    setLevels(next)
    setSelectedIndex(next.length - 1)
  }

  const handleDeleteLevel = () => {
    if (levels.length <= 1) return
    const next = levels.filter((_, i) => i !== selectedIndex)
    setLevels(next)
    setSelectedIndex(Math.max(0, selectedIndex - 1))
  }

  const handleCopyLevel = () => {
    const cur = levels[selectedIndex]
    if (!cur) return
    const copy = { ...cur }
    const next = [...levels.slice(0, selectedIndex + 1), copy, ...levels.slice(selectedIndex + 1)]
    setLevels(next)
    setSelectedIndex(selectedIndex + 1)
  }

  const handleMoveUp = () => {
    if (selectedIndex <= 0) return
    const next = [...levels]
    const temp = next[selectedIndex]
    next[selectedIndex] = next[selectedIndex - 1]
    next[selectedIndex - 1] = temp
    setLevels(next)
    setSelectedIndex(selectedIndex - 1)
  }

  const handleMoveDown = () => {
    if (selectedIndex >= levels.length - 1) return
    const next = [...levels]
    const temp = next[selectedIndex]
    next[selectedIndex] = next[selectedIndex + 1]
    next[selectedIndex + 1] = temp
    setLevels(next)
    setSelectedIndex(selectedIndex + 1)
  }

  const updateLevel = (index: number, patch: Partial<SortLevel>) => {
    setLevels((prev) => prev.map((lvl, i) => (i === index ? { ...lvl, ...patch } : lvl)))
  }

  const handleOk = () => {
    onClose()
    runMaybeBusy((range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1), 'Sorting...', () => sortRange(store, sheet, range, {
      hasHeaders,
      caseSensitive,
      levels
    }))
  }

  return (
    <>
      <Dialog title="Sort" width={560} onOk={handleOk} onCancel={onClose}>
        <div className="xl-sort-body">
          <div className="xl-sort-toolbar">
            <div className="xl-sort-toolbar-left">
              <button className="xl-sort-btn" onClick={handleAddLevel}>
                + Add Level
              </button>
              <button
                className="xl-sort-btn"
                disabled={levels.length <= 1}
                onClick={handleDeleteLevel}
              >
                Delete Level
              </button>
              <button className="xl-sort-btn" onClick={handleCopyLevel}>
                Copy Level
              </button>
              <button
                className="xl-sort-btn xl-sort-icon-btn"
                disabled={selectedIndex <= 0}
                title="Move Up"
                onClick={handleMoveUp}
              >
                ▲
              </button>
              <button
                className="xl-sort-btn xl-sort-icon-btn"
                disabled={selectedIndex >= levels.length - 1}
                title="Move Down"
                onClick={handleMoveDown}
              >
                ▼
              </button>
              <button className="xl-sort-btn" onClick={() => setShowOptions(true)}>
                Options...
              </button>
            </div>

            <label className="xl-sort-headers-checkbox">
              <input
                type="checkbox"
                checked={hasHeaders}
                onChange={(e) => setHasHeaders(e.target.checked)}
              />
              <span>My data has headers</span>
            </label>
          </div>

          <div className="xl-sort-table-wrapper">
            <table className="xl-sort-table">
              <thead>
                <tr>
                  <th style={{ width: '35%' }}>Column</th>
                  <th style={{ width: '30%' }}>Sort On</th>
                  <th style={{ width: '35%' }}>Order</th>
                </tr>
              </thead>
              <tbody>
                {levels.map((lvl, idx) => {
                  const isFirst = idx === 0
                  const isSelected = idx === selectedIndex
                  return (
                    <tr
                      key={idx}
                      className={isSelected ? 'selected' : ''}
                      onClick={() => setSelectedIndex(idx)}
                    >
                      <td>
                        <span className="xl-sort-row-prefix">{isFirst ? 'Sort by' : 'Then by'}</span>
                        <select
                          className="xl-sort-select"
                          value={lvl.columnIndex}
                          onChange={(e) => updateLevel(idx, { columnIndex: +e.target.value })}
                        >
                          {columns.map((col) => (
                            <option key={col.index} value={col.index}>
                              {col.name}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>
                        <select
                          className="xl-sort-select"
                          value={lvl.sortOn}
                          onChange={(e) => updateLevel(idx, { sortOn: e.target.value as any })}
                        >
                          <option value="values">Cell Values</option>
                          <option value="cellColor">Cell Color</option>
                          <option value="fontColor">Font Color</option>
                        </select>
                      </td>
                      <td>
                        <select
                          className="xl-sort-select"
                          value={lvl.order}
                          onChange={(e) => updateLevel(idx, { order: e.target.value as any })}
                        >
                          <option value="asc">A to Z / Smallest to Largest</option>
                          <option value="desc">Z to A / Largest to Smallest</option>
                        </select>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      </Dialog>

      {showOptions && (
        <Dialog
          title="Sort Options"
          width={320}
          onOk={() => setShowOptions(false)}
          onCancel={() => setShowOptions(false)}
        >
          <div className="xl-sort-options-body">
            <label className="xl-sort-radio-label">
              <input
                type="checkbox"
                checked={caseSensitive}
                onChange={(e) => setCaseSensitive(e.target.checked)}
              />
              <span>Case sensitive</span>
            </label>
            <div className="xl-sort-orient-group">
              <span className="xl-sort-orient-title">Orientation</span>
              <label className="xl-sort-radio-label">
                <input type="radio" checked readOnly name="sortOrientation" />
                <span>Sort top to bottom</span>
              </label>
              <label className="xl-sort-radio-label">
                <input type="radio" disabled name="sortOrientation" />
                <span>Sort left to right</span>
              </label>
            </div>
          </div>
        </Dialog>
      )}
    </>
  )
}

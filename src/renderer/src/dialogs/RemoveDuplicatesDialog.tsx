import { runMaybeBusy } from '../chrome/ui'
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { colName, normRange } from '../model/types'
import { detectHeaders } from '../data/sort'
import { removeDuplicates } from '../data/removeDuplicates'
import './RemoveDuplicatesDialog.css'

export interface RemoveDuplicatesDialogProps {
  store: WorkbookStore
  onClose: () => void
}

export function RemoveDuplicatesDialog({ store, onClose }: RemoveDuplicatesDialogProps) {
  const sheet = store.active
  // A single active cell means "this table": Excel expands to the current region and selects it.
  const [range] = useState(() => {
    const sel = normRange(sheet.selection.ranges[0] || { r1: 0, c1: 0, r2: 0, c2: 0 })
    if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) return sel
    const region = normRange(store.currentRegion(sheet, sel.r1, sel.c1))
    store.setSelection({ ranges: [region], active: { r: sel.r1, c: sel.c1 } })
    return region
  })

  const [hasHeaders, setHasHeaders] = useState<boolean>(() => detectHeaders(store, sheet, range))

  const columns: { index: number; name: string }[] = []
  for (let c = range.c1; c <= range.c2; c++) {
    let header = ''
    if (hasHeaders) {
      header = store.getRaw(sheet, range.r1, c).trim()
    }
    columns.push({ index: c, name: header || 'Column ' + colName(c) })
  }

  const [selectedCols, setSelectedCols] = useState<number[]>(() => columns.map((c) => c.index))
  const [resultMsg, setResultMsg] = useState<string | null>(null)

  const handleSelectAll = () => {
    setSelectedCols(columns.map((c) => c.index))
  }

  const handleUnselectAll = () => {
    setSelectedCols([])
  }

  const toggleCol = (index: number) => {
    setSelectedCols((prev) =>
      prev.includes(index) ? prev.filter((i) => i !== index) : [...prev, index]
    )
  }

  const handleOk = () => {
    if (selectedCols.length === 0) return
    const run = (): void => {
      const res = removeDuplicates(store, sheet, range, selectedCols, hasHeaders)
      setResultMsg(`${res.removedCount} duplicate values found and removed; ${res.remainingCount} unique values remain.`)
    }
    runMaybeBusy((range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1), 'Removing duplicates...', run)
  }

  return (
    <>
      <Dialog title="Remove Duplicates" width={440} onOk={handleOk} onCancel={onClose}>
        <div className="xl-rd-body">
          <p className="xl-rd-prompt">
            To delete duplicate values, select one or more columns that contain duplicates.
          </p>

          <div className="xl-rd-toolbar">
            <div className="xl-rd-btn-group">
              <button className="xl-rd-btn" onClick={handleSelectAll}>
                Select All
              </button>
              <button className="xl-rd-btn" onClick={handleUnselectAll}>
                Unselect All
              </button>
            </div>

            <label className="xl-rd-headers-checkbox">
              <input
                type="checkbox"
                checked={hasHeaders}
                onChange={(e) => setHasHeaders(e.target.checked)}
              />
              <span>My data has headers</span>
            </label>
          </div>

          <div className="xl-rd-columns-box">
            <div className="xl-rd-columns-header">Columns</div>
            <div className="xl-rd-columns-list">
              {columns.map((col) => (
                <label key={col.index} className="xl-rd-column-item">
                  <input
                    type="checkbox"
                    checked={selectedCols.includes(col.index)}
                    onChange={() => toggleCol(col.index)}
                  />
                  <span>{col.name}</span>
                </label>
              ))}
            </div>
          </div>
        </div>
      </Dialog>

      {resultMsg && (
        <Dialog
          title="Microsoft Excel"
          width={380}
          onOk={onClose}
          onCancel={onClose}
        >
          <div className="xl-rd-result-body">
            <p>{resultMsg}</p>
          </div>
        </Dialog>
      )}
    </>
  )
}

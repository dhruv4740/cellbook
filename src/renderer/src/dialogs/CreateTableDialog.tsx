// Home > Format as Table / Insert > Table / Ctrl+T / Ctrl+L: "Create Table" dialog.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { Range, parseA1, rangeA1 } from '../model/types'
import { createTable } from '../model/tables'
import { detectHeaders } from '../data/sort'
import { setUi } from '../chrome/ui'
import './RemoveDuplicatesDialog.css'
import './TableDialogs.css'

export function CreateTableDialog({ store, range, styleName, onClose }: { store: WorkbookStore; range: Range; styleName?: string; onClose: () => void }) {
  const sheet = store.active
  const [refText, setRefText] = useState(rangeA1(range))
  const [hasHeaders, setHasHeaders] = useState(() => detectHeaders(store, sheet, range))
  const [error, setError] = useState<string | null>(null)

  const handleOk = () => {
    const rg = parseA1(refText)
    if (!rg) { setError('The reference is not valid.') ; return }
    createTable(store, sheet, rg, hasHeaders, styleName)
    setUi({ ribbonTab: 'Table Design' })
    onClose()
  }

  return (
    <Dialog title="Create Table" width={380} onOk={handleOk} onCancel={onClose}>
      <div className="xl-rd-body">
        <p className="xl-rd-prompt">Where is the data for your table?</p>
        <input
          className="xl-ct-ref"
          value={refText}
          onChange={(e) => { setRefText(e.target.value); setError(null) }}
          autoFocus
        />
        {error && <p className="xl-rd-prompt" style={{ color: '#c00' }}>{error}</p>}
        <label className="xl-rd-headers-checkbox">
          <input type="checkbox" checked={hasHeaders} onChange={(e) => setHasHeaders(e.target.checked)} />
          <span>My table has headers</span>
        </label>
      </div>
    </Dialog>
  )
}

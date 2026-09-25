// Table Design > Resize Table.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { Table, parseA1, rangeA1 } from '../model/types'
import { resizeTable } from '../model/tables'
import './RemoveDuplicatesDialog.css'
import './TableDialogs.css'

export function ResizeTableDialog({ store, table, onClose }: { store: WorkbookStore; table: Table; onClose: () => void }) {
  const [refText, setRefText] = useState(rangeA1(table.range))
  const [error, setError] = useState<string | null>(null)

  const handleOk = () => {
    const rg = parseA1(refText)
    if (!rg) { setError('The reference is not valid.'); return }
    if (rg.c1 === rg.c2 && rg.r1 === rg.r2) { setError('The Table must overlap the original Table range.'); return }
    resizeTable(store, store.active, table, rg)
    onClose()
  }

  return (
    <Dialog title="Resize Table" width={380} onOk={handleOk} onCancel={onClose}>
      <div className="xl-rd-body">
        <p className="xl-rd-prompt">Where is the data for your table?</p>
        <input className="xl-ct-ref" value={refText} onChange={(e) => { setRefText(e.target.value); setError(null) }} autoFocus />
        {error && <p className="xl-rd-prompt" style={{ color: '#c00' }}>{error}</p>}
      </div>
    </Dialog>
  )
}

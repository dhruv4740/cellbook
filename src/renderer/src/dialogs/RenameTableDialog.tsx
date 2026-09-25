import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { Table } from '../model/types'
import { renameTable } from '../model/tables'
import './TableDialogs.css'

export function RenameTableDialog({ store, table, onClose }: { store: WorkbookStore; table: Table; onClose: () => void }) {
  const [name, setName] = useState(table.name)
  const [error, setError] = useState<string | null>(null)
  const handleOk = () => {
    const err = renameTable(store, store.active, table, name)
    if (err) { setError(err); return }
    onClose()
  }
  return (
    <Dialog title="Table Name" width={340} onOk={handleOk} onCancel={onClose}>
      <div style={{ font: '12px var(--ui-font, "Segoe UI", Tahoma, sans-serif)', display: 'flex', flexDirection: 'column', gap: 8 }}>
        <label>Table name:</label>
        <input className="xl-ct-ref" value={name} onChange={(e) => { setName(e.target.value); setError(null) }} autoFocus />
        {error && <p style={{ color: '#c00', margin: 0 }}>{error}</p>}
      </div>
    </Dialog>
  )
}

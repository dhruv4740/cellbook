// View > Custom Views: Add / Show / Delete named views (session only).
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { addView, deleteView, listViews, showView } from '../data/customViews'
import { store } from '../model/store'

export function CustomViewsDialog({ onClose }: { onClose: () => void }) {
  const [, bump] = useState(0)
  const [sel, setSel] = useState('')
  const [name, setName] = useState('')
  const list = listViews()
  const sheetName = (uid: number): string => store.sheets.find((s) => s.uid === uid)?.name ?? '(deleted sheet)'
  return (
    <Dialog title="Custom Views" width={400} onCancel={onClose} cancelLabel="Close" onOk={undefined}>
      <div style={{ fontSize: 12 }}>
        <div>Views (this session):</div>
        <select size={6} style={{ width: '100%', margin: '4px 0' }} value={sel} onChange={(e) => setSel(e.target.value)} onDoubleClick={() => { const v = list.find((x) => x.name === sel); if (v && showView(v)) onClose() }}>
          {list.map((v) => <option key={v.name} value={v.name}>{v.name} - {sheetName(v.sheetUid)}, {v.zoom}%</option>)}
        </select>
        <div className="xl-field-row"><label style={{ width: 60 }}>Name:</label><input style={{ flex: 1 }} value={name} onChange={(e) => setName(e.target.value)} /></div>
        <div style={{ marginTop: 8, display: 'flex', gap: 6 }}>
          <button className="xl-btn" disabled={!name.trim()} onClick={() => { addView(name.trim()); setSel(name.trim()); setName(''); bump((n) => n + 1) }}>Add current view</button>
          <button className="xl-btn" disabled={!sel} onClick={() => { const v = list.find((x) => x.name === sel); if (v && showView(v)) onClose() }}>Show</button>
          <button className="xl-btn" disabled={!sel} onClick={() => { deleteView(sel); setSel(''); bump((n) => n + 1) }}>Delete</button>
        </div>
      </div>
    </Dialog>
  )
}

// Workbooks opened here carry no VBA, queries or connections, so these dialogs show Excel's empty-workbook state.
import { Dialog } from '../components/Dialog'

export function MacrosDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog title="Macro" width={420} onCancel={onClose} cancelLabel="Cancel" onOk={undefined}>
      <div style={{ fontSize: 12 }}>
        <div className="xl-field-row"><label style={{ width: 90 }}>Macro name:</label><input style={{ flex: 1 }} disabled /></div>
        <select size={7} style={{ width: '100%', margin: '6px 0' }} disabled />
        <div style={{ color: '#605e5c', marginBottom: 6 }}>No macros in this workbook. Cellbook opens and saves .xlsx files, which cannot contain VBA.</div>
        <div style={{ display: 'flex', gap: 6 }}>{['Run', 'Step Into', 'Edit', 'Create', 'Delete'].map((b) => <button key={b} className="xl-btn" disabled>{b}</button>)}</div>
      </div>
    </Dialog>
  )
}

export function ConnectionsDialog({ title, message, onClose }: { title: string; message: string; onClose: () => void }) {
  return (
    <Dialog title={title} width={420} onCancel={onClose} cancelLabel="Close" onOk={undefined}>
      <div style={{ fontSize: 12, minHeight: 80 }}>{message}</div>
    </Dialog>
  )
}

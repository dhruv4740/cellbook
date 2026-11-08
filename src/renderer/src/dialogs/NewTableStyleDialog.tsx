// Home > Format as Table > New Table Style: pick the header, stripe and border colours; the style joins the Custom gallery.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { TableStyleDef, loadCustomTableStyles, saveCustomTableStyles, upsertStyle } from '../format/customTableStyles'
import { formatAsTable } from '../grid/commands'

export function NewTableStyleDialog({ onClose }: { onClose: () => void }) {
  const [name, setName] = useState(() => `Table Style ${loadCustomTableStyles().length + 1}`)
  const [header, setHeader] = useState('#4472c4')
  const [headerFont, setHeaderFont] = useState('#ffffff')
  const [band, setBand] = useState('#d9e1f2')
  const [line, setLine] = useState('#8ea9db')
  const [apply, setApply] = useState(true)
  const style: TableStyleDef = { name: name.trim(), header, band, line, headerFont }
  const ok = (): void => {
    if (!style.name) return
    saveCustomTableStyles(upsertStyle(loadCustomTableStyles(), style))
    onClose()
    if (apply) formatAsTable(style)
  }
  const row = (label: string, v: string, set: (s: string) => void): React.ReactElement => (
    <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 130 }}>{label}</label><input type="color" value={v} onChange={(e) => set(e.target.value)} /></div>
  )
  return (
    <Dialog title="New Table Style" width={380} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <div className="xl-field-row"><label style={{ width: 130 }}>Name:</label><input autoFocus style={{ flex: 1 }} value={name} onChange={(e) => setName(e.target.value)} /></div>
        {row('Header row fill:', header, setHeader)}
        {row('Header row text:', headerFont, setHeaderFont)}
        {row('First row stripe:', band, setBand)}
        {row('Whole table border:', line, setLine)}
        <div style={{ margin: '10px 0 4px', width: 160, border: `1px solid ${line}` }}>
          <div style={{ background: header, color: headerFont, fontWeight: 700, padding: '1px 4px' }}>Header</div>
          <div style={{ background: '#fff', padding: '1px 4px' }}>Row 1</div>
          <div style={{ background: band, padding: '1px 4px' }}>Row 2</div>
        </div>
        <label><input type="checkbox" checked={apply} onChange={(e) => setApply(e.target.checked)} /> Apply to the selected cells now</label>
      </div>
    </Dialog>
  )
}

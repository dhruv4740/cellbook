// Review > Allow Edit Ranges: named ranges that stay editable once the sheet is protected.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store, useStore } from '../model/store'
import { EditRange } from '../model/types'
import { parseSqref } from '../model/protection'
import { rangeA1, normRange } from '../model/types'

export function EditRangesDialog({ onClose }: { onClose: () => void }) {
  useStore()
  const sheet = store.active
  const list = sheet.layout.editRanges ?? []
  const [sel, setSel] = useState(-1)
  const [title, setTitle] = useState('')
  const [ref, setRef] = useState(() => rangeA1(normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])))
  const [err, setErr] = useState<string | null>(null)
  const save = (next: EditRange[]): void => store.updateLayout(sheet, 'Allow Edit Ranges', (l) => { l.editRanges = next.length ? next : undefined })
  const pick = (i: number): void => { setSel(i); setTitle(list[i].title); setRef(list[i].ref); setErr(null) }
  const entry = (): EditRange | null => {
    const t = title.trim() || `Range${list.length + 1}`
    const parts = parseSqref(ref)
    if (!parts.length) { setErr('Enter a valid range, such as A1:B4.'); return null }
    if (list.some((e, i) => i !== sel && e.title.toLowerCase() === t.toLowerCase())) { setErr('A range with that title already exists.'); return null }
    setErr(null)
    return { title: t, ref: parts.map(rangeA1).join(' ') }
  }
  const add = (): void => { const e = entry(); if (e) { save([...list, e]); setSel(list.length); setTitle(e.title) } }
  const modify = (): void => { const e = entry(); if (e && sel >= 0) save(list.map((x, i) => (i === sel ? e : x))) }
  const del = (): void => { if (sel >= 0) { save(list.filter((_, i) => i !== sel)); setSel(-1); setTitle('') } }
  return (
    <Dialog title="Allow Users to Edit Ranges" width={440} onCancel={onClose} cancelLabel="Close" onOk={undefined}>
      <div style={{ fontSize: 12 }}>
        <div>Ranges that stay editable when the sheet is protected:</div>
        <select size={6} style={{ width: '100%', margin: '4px 0' }} value={sel} onChange={(e) => pick(+e.target.value)}>
          {list.map((e, i) => <option key={e.title} value={i}>{e.title} - {e.ref}</option>)}
        </select>
        <div className="xl-field-row"><label style={{ width: 80 }}>Title:</label><input style={{ flex: 1 }} value={title} onChange={(e) => setTitle(e.target.value)} /></div>
        <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 80 }}>Refers to:</label><input style={{ flex: 1 }} value={ref} onChange={(e) => setRef(e.target.value)} /></div>
        <div style={{ marginTop: 8, display: 'flex', gap: 6 }}>
          <button className="xl-btn" onClick={add}>New</button>
          <button className="xl-btn" disabled={sel < 0} onClick={modify}>Modify</button>
          <button className="xl-btn" disabled={sel < 0} onClick={del}>Delete</button>
        </div>
        {err && <div style={{ color: '#a4262c', marginTop: 6 }}>{err}</div>}
      </div>
    </Dialog>
  )
}

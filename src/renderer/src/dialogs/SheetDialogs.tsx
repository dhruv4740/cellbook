// Sheet-level dialogs: Unhide, Move or Copy, and a plain Excel-style message box.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { fillSeries } from '../grid/commands'
import { SeriesType, DateUnit } from '../engine/series'

export function MessageDialog({ title = 'Microsoft Excel', text, onClose }: { title?: string; text: string; onClose: () => void }) {
  return (
    <Dialog title={title} onCancel={onClose} cancelLabel="OK" width={360}>
      <div style={{ whiteSpace: 'pre-wrap', fontSize: 12 }}>{text}</div>
    </Dialog>
  )
}

export function UnhideSheetDialog({ onClose }: { onClose: () => void }) {
  const hidden = store.sheets.map((s, i) => ({ s, i })).filter((x) => x.s.layout.sheetHidden === 'hidden')
  const [sel, setSel] = useState(hidden[0]?.i ?? -1)
  if (!hidden.length) return <MessageDialog text="No sheets are hidden." onClose={onClose} />
  return (
    <Dialog title="Unhide" width={320} onCancel={onClose} onOk={() => { if (sel >= 0) store.unhideSheet(sel); onClose() }}>
      <div style={{ fontSize: 12, marginBottom: 4 }}>Unhide sheet:</div>
      <select size={8} style={{ width: '100%' }} value={sel} onChange={(e) => setSel(+e.target.value)} onDoubleClick={() => { store.unhideSheet(sel); onClose() }}>
        {hidden.map((x) => <option key={x.s.uid} value={x.i}>{x.s.name}</option>)}
      </select>
    </Dialog>
  )
}

export function MoveCopySheetDialog({ index, onClose }: { index: number; onClose: () => void }) {
  const [before, setBefore] = useState(Math.min(index + 1, store.sheets.length))
  const [copy, setCopy] = useState(false)
  const ok = () => {
    if (copy) store.copySheet(index, before)
    else {
      const to = before > index ? before - 1 : before
      store.moveSheet(index, to)
    }
    onClose()
  }
  return (
    <Dialog title="Move or Copy" width={320} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12, marginBottom: 4 }}>Move selected sheets<br />Before sheet:</div>
      <select size={8} style={{ width: '100%' }} value={before} onChange={(e) => setBefore(+e.target.value)}>
        {store.sheets.map((s, i) => <option key={s.uid} value={i}>{s.name}</option>)}
        <option value={store.sheets.length}>(move to end)</option>
      </select>
      <label style={{ display: 'block', marginTop: 8, fontSize: 12 }}>
        <input type="checkbox" checked={copy} onChange={(e) => setCopy(e.target.checked)} /> Create a copy
      </label>
    </Dialog>
  )
}

export function SeriesDialog({ onClose }: { onClose: () => void }) {
  const n = store.active.selection.ranges[store.active.selection.ranges.length - 1]
  const wide = Math.abs(n.c2 - n.c1) > Math.abs(n.r2 - n.r1)
  const [byRows, setByRows] = useState(wide)
  const [type, setType] = useState<SeriesType>('linear')
  const [unit, setUnit] = useState<DateUnit>('day')
  const [step, setStep] = useState('1')
  const [stop, setStop] = useState('')
  const [err, setErr] = useState<string | null>(null)
  const ok = () => {
    const st = Number(step)
    if (!isFinite(st)) { setErr('Step value must be a number.'); return }
    const sp = stop.trim() === '' ? undefined : Number(stop)
    if (sp !== undefined && !isFinite(sp)) { setErr('Stop value must be a number.'); return }
    const e = fillSeries({ byRows, type, unit, step: st, stop: sp })
    if (e) setErr(e); else onClose()
  }
  const radio = (name: string, checked: boolean, onChange: () => void, label: string, disabled = false) => (
    <label style={{ display: 'block', opacity: disabled ? 0.5 : 1 }}><input type="radio" name={name} checked={checked} disabled={disabled} onChange={onChange} /> {label}</label>
  )
  return (
    <Dialog title="Series" width={380} onCancel={onClose} onOk={ok}>
      <div style={{ display: 'flex', gap: 16, fontSize: 12 }}>
        <fieldset><legend>Series in</legend>{radio('sin', byRows, () => setByRows(true), 'Rows')}{radio('sin', !byRows, () => setByRows(false), 'Columns')}</fieldset>
        <fieldset><legend>Type</legend>{radio('st', type === 'linear', () => setType('linear'), 'Linear')}{radio('st', type === 'growth', () => setType('growth'), 'Growth')}{radio('st', type === 'date', () => setType('date'), 'Date')}</fieldset>
        <fieldset><legend>Date unit</legend>
          {(['day', 'weekday', 'month', 'year'] as DateUnit[]).map((u) => <span key={u}>{radio('du', unit === u, () => setUnit(u), u === 'weekday' ? 'Weekday' : u[0].toUpperCase() + u.slice(1), type !== 'date')}</span>)}
        </fieldset>
      </div>
      <div style={{ display: 'flex', gap: 12, marginTop: 10, fontSize: 12 }}>
        <label>Step value: <input autoFocus value={step} onChange={(e) => setStep(e.target.value)} style={{ width: 70 }} /></label>
        <label>Stop value: <input value={stop} onChange={(e) => setStop(e.target.value)} style={{ width: 70 }} /></label>
      </div>
      {err && <div style={{ color: '#a4262c', fontSize: 12, marginTop: 8 }}>{err}</div>}
    </Dialog>
  )
}

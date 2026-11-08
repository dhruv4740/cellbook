// Formulas > Define Name and Name Manager.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store, useStore } from '../model/store'
import { colName, normRange } from '../model/types'
import { applicableNames, applyNames } from '../data/applyNames'

function quoteSheet(n: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`
}
/** The selection as an absolute reference, e.g. =Sheet1!$A$2:$A$20 (what Excel pre-fills). */
export function selectionRef(): string {
  const sh = store.active
  const n = normRange(sh.selection.ranges[sh.selection.ranges.length - 1])
  const a = (r: number, c: number) => `$${colName(c)}$${r + 1}`
  return `=${quoteSheet(sh.name)}!${a(n.r1, n.c1)}${n.r1 === n.r2 && n.c1 === n.c2 ? '' : ':' + a(n.r2, n.c2)}`
}

/** Suggested name: the text in the cell above / left of the selection, like Excel. */
function suggestName(): string {
  const sh = store.active
  const n = normRange(sh.selection.ranges[sh.selection.ranges.length - 1])
  const t = String(store.getRaw(sh, n.r1, n.c1) || (n.r1 > 0 ? store.getRaw(sh, n.r1 - 1, n.c1) : '') || '')
  const s = t.trim().replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^([0-9])/, '_$1')
  return /^[A-Za-z_]/.test(s) && !t.startsWith('=') ? s.slice(0, 255) : ''
}

export function DefineNameDialog({ onClose, edit }: { onClose: () => void; edit?: { name: string; sheetUid?: number } }) {
  useStore()
  const [name, setName] = useState(edit?.name ?? suggestName())
  const [scope, setScope] = useState<number | undefined>(edit?.sheetUid)
  const [ref, setRef] = useState(edit ? store.nameRef(edit.name, edit.sheetUid) : selectionRef())
  const [err, setErr] = useState<string | null>(null)
  const ok = () => {
    const e = store.defineName(name, ref, scope, edit)
    if (e) setErr(e)
    else onClose()
  }
  return (
    <Dialog title={edit ? 'Edit Name' : 'New Name'} width={420} onCancel={onClose} onOk={ok}>
      <div className="xl-field-row"><label style={{ width: 80 }}>Name:</label><input autoFocus value={name} onChange={(e) => setName(e.target.value)} style={{ flex: 1 }} /></div>
      <div className="xl-field-row" style={{ marginTop: 6 }}>
        <label style={{ width: 80 }}>Scope:</label>
        <select value={scope ?? ''} onChange={(e) => setScope(e.target.value === '' ? undefined : +e.target.value)} disabled={!!edit}>
          <option value="">Workbook</option>
          {store.sheets.map((s) => <option key={s.uid} value={s.uid}>{s.name}</option>)}
        </select>
      </div>
      <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 80 }}>Refers to:</label><input value={ref} onChange={(e) => setRef(e.target.value)} style={{ flex: 1 }} /></div>
      {err && <div style={{ color: '#a4262c', fontSize: 12, marginTop: 8, whiteSpace: 'pre-wrap' }}>{err}</div>}
    </Dialog>
  )
}

export function NameManagerDialog({ onClose }: { onClose: () => void }) {
  useStore()
  const list = store.names.filter((n) => !/^_xlnm\./i.test(n.name) && !n.hidden)
  const [sel, setSel] = useState(0)
  const [editing, setEditing] = useState<null | 'new' | { name: string; sheetUid?: number }>(null)
  const cur = list[Math.min(sel, list.length - 1)]
  if (editing) return <DefineNameDialog onClose={() => setEditing(null)} edit={editing === 'new' ? undefined : editing} />
  const valueOf = (n: (typeof list)[number]): string => {
    try {
      const sh = n.sheetUid !== undefined ? store.sheetByUid(n.sheetUid) : undefined
      const v = store.hf.getNamedExpressionValue(n.name, sh ? store.hfId(sh) : undefined)
      if (Array.isArray(v)) return '{' + (v as unknown[][]).slice(0, 3).map((row) => row.slice(0, 3).map((x: unknown) => (typeof x === 'object' && x ? (x as { value?: string }).value ?? '' : String(x ?? ''))).join(',')).join(';') + (v.length > 3 ? ';…' : '') + '}'
      if (v && typeof v === 'object') return String((v as { value?: string }).value ?? '#ERROR!')
      return String(v ?? '')
    } catch { return '' }
  }
  return (
    <Dialog title="Name Manager" width={640} onCancel={onClose} cancelLabel="Close">
      <div style={{ display: 'flex', gap: 6, marginBottom: 8 }}>
        <button className="xl-btn" onClick={() => setEditing('new')}>New...</button>
        <button className="xl-btn" disabled={!cur} onClick={() => cur && setEditing({ name: cur.name, sheetUid: cur.sheetUid })}>Edit...</button>
        <button className="xl-btn" disabled={!cur} onClick={() => cur && store.deleteName(cur.name, cur.sheetUid)}>Delete</button>
      </div>
      <div style={{ border: '1px solid #c8c8c8', height: 220, overflow: 'auto', fontSize: 12 }}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead><tr style={{ background: '#f3f3f3', textAlign: 'left' }}><th>Name</th><th>Value</th><th>Refers To</th><th>Scope</th></tr></thead>
          <tbody>
            {list.map((n, i) => (
              <tr key={n.name + ':' + n.sheetUid} onClick={() => setSel(i)} onDoubleClick={() => setEditing({ name: n.name, sheetUid: n.sheetUid })} style={{ background: i === sel ? '#cfe3f7' : undefined, cursor: 'default' }}>
                <td>{n.name}</td><td>{valueOf(n)}</td><td>{store.nameRef(n.name, n.sheetUid)}</td>
                <td>{n.sheetUid !== undefined ? store.sheetByUid(n.sheetUid)?.name : 'Workbook'}</td>
              </tr>
            ))}
            {!list.length && <tr><td colSpan={4} style={{ color: '#777', padding: 8 }}>No names defined. Click New... to create one.</td></tr>}
          </tbody>
        </table>
      </div>
    </Dialog>
  )
}

/** Formulas > Create from Selection: names each column (or row) of the selection after its header. */
export function CreateFromSelectionDialog({ onClose }: { onClose: () => void }) {
  const [top, setTop] = useState(true)
  const [left, setLeft] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const ok = () => {
    const sh = store.active
    const n = normRange(sh.selection.ranges[sh.selection.ranges.length - 1])
    const q = quoteSheet(sh.name)
    const a = (r: number, c: number) => `$${colName(c)}$${r + 1}`
    const fails: string[] = []
    store.transact('Create Names', () => {
      if (top) for (let c = n.c1 + (left ? 1 : 0); c <= n.c2; c++) {
        const nm = String(store.getDisplay(sh, n.r1, c).text).trim().replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^([0-9])/, '_$1')
        if (!nm || n.r2 <= n.r1) continue
        const e = store.defineName(nm, `=${q}!${a(n.r1 + 1, c)}:${a(n.r2, c)}`)
        if (e) fails.push(nm)
      }
      if (left) for (let r = n.r1 + (top ? 1 : 0); r <= n.r2; r++) {
        const nm = String(store.getDisplay(sh, r, n.c1).text).trim().replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^([0-9])/, '_$1')
        if (!nm || n.c2 <= n.c1) continue
        const e = store.defineName(nm, `=${q}!${a(r, n.c1 + 1)}:${a(r, n.c2)}`)
        if (e) fails.push(nm)
      }
    })
    if (fails.length) setErr('These names were not created (invalid or already used): ' + fails.join(', '))
    else onClose()
  }
  return (
    <Dialog title="Create Names from Selection" width={320} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>Create names from values in the:</div>
      <label style={{ display: 'block' }}><input type="checkbox" checked={top} onChange={(e) => setTop(e.target.checked)} /> Top row</label>
      <label style={{ display: 'block' }}><input type="checkbox" checked={left} onChange={(e) => setLeft(e.target.checked)} /> Left column</label>
      {err && <div style={{ color: '#a4262c', fontSize: 12, marginTop: 8 }}>{err}</div>}
    </Dialog>
  )
}

/** Formulas > Define Name > Apply Names: pick the names to substitute into formulas. */
export function ApplyNamesDialog({ onClose }: { onClose: () => void }) {
  const list = applicableNames()
  const [picked, setPicked] = useState<string[]>([])
  const [msg, setMsg] = useState('')
  const ok = () => {
    if (!picked.length) return onClose()
    const n = applyNames(store.active, picked)
    if (n) onClose(); else setMsg('No references to those names were found in this sheet’s formulas.')
  }
  return (
    <Dialog title="Apply Names" width={320} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <div>Apply names:</div>
        <select multiple size={8} style={{ width: '100%' }} value={picked} onChange={(e) => setPicked(Array.from(e.target.selectedOptions).map((o) => o.value))}>
          {list.map((n) => <option key={n.name} value={n.name}>{n.name}</option>)}
        </select>
        {!list.length && <div style={{ color: '#666', marginTop: 6 }}>There are no names that refer to a cell or range.</div>}
        {msg && <div style={{ color: '#a4262c', marginTop: 6 }}>{msg}</div>}
      </div>
    </Dialog>
  )
}

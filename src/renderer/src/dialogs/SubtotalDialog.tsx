// Data > Subtotal.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { normRange, colName } from '../model/types'
import { applySubtotals, removeSubtotals, SUBTOTAL_FUNCS } from '../data/subtotal'

export function SubtotalDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const n = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
  const rg = n.r1 !== n.r2 || n.c1 !== n.c2 ? store.clampToUsed(sheet, n) : store.currentRegion(sheet, sheet.selection.active.r, sheet.selection.active.c)
  const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i)
  const head = (c: number) => store.getDisplay(sheet, rg.r1, c).text || `Column ${colName(c)}`
  const [groupCol, setGroupCol] = useState(cols[0])
  const [func, setFunc] = useState(0)
  const lastNumeric = [...cols].reverse().find((c) => typeof store.getValue(sheet, rg.r1 + 1, c) === 'number') ?? cols[cols.length - 1]
  const [addTo, setAddTo] = useState<number[]>([lastNumeric])
  const [replace, setReplace] = useState(true)
  const [below, setBelow] = useState(true)
  const [err, setErr] = useState<string | null>(null)
  if (rg.r2 <= rg.r1) {
    return (
      <Dialog title="Microsoft Excel" onCancel={onClose} cancelLabel="OK" width={380}>
        <div style={{ fontSize: 12 }}>Select a cell in a list with a header row and data, then choose Subtotal.</div>
      </Dialog>
    )
  }
  const ok = () => {
    const e = applySubtotals(store, sheet, rg, { groupCol, func, addTo, replace, summaryBelow: below })
    if (e) setErr(e); else onClose()
  }
  return (
    <Dialog title="Subtotal" width={340} onCancel={onClose} onOk={ok}
      footerLeft={<button className="xl-btn" onClick={() => { removeSubtotals(store, sheet, rg); onClose() }}>Remove All</button>}>
      <div style={{ fontSize: 12 }}>
        <div>At each change in:</div>
        <select style={{ width: '100%' }} value={groupCol} onChange={(e) => setGroupCol(+e.target.value)}>
          {cols.map((c) => <option key={c} value={c}>{head(c)}</option>)}
        </select>
        <div style={{ marginTop: 8 }}>Use function:</div>
        <select style={{ width: '100%' }} value={func} onChange={(e) => setFunc(+e.target.value)}>
          {SUBTOTAL_FUNCS.map((f, i) => <option key={f.name} value={i}>{f.name}</option>)}
        </select>
        <div style={{ marginTop: 8 }}>Add subtotal to:</div>
        <div style={{ border: '1px solid #c8c8c8', height: 110, overflow: 'auto', padding: 2 }}>
          {cols.map((c) => (
            <label key={c} style={{ display: 'block' }}>
              <input type="checkbox" checked={addTo.includes(c)} onChange={(e) => setAddTo(e.target.checked ? [...addTo, c] : addTo.filter((x) => x !== c))} /> {head(c)}
            </label>
          ))}
        </div>
        <label style={{ display: 'block', marginTop: 8 }}><input type="checkbox" checked={replace} onChange={(e) => setReplace(e.target.checked)} /> Replace current subtotals</label>
        <label style={{ display: 'block' }}><input type="checkbox" checked={below} onChange={(e) => setBelow(e.target.checked)} /> Summary below data</label>
        <div style={{ color: '#666', marginTop: 6 }}>Sort the list by the "At each change in" column first, so each group's rows are together.</div>
        {err && <div style={{ color: '#a4262c', marginTop: 6 }}>{err}</div>}
      </div>
    </Dialog>
  )
}

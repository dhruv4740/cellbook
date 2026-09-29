// Data > What-If Analysis > Data Table and Scenario Manager.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { colName, rangeA1, ScenarioDef } from '../model/types'
import { computeDataTable, scenariosOf, saveScenario, deleteScenario, showScenario, scenarioSummary, cellsOfRanges } from '../data/whatIf'
import { parseRef } from './DataToolDialogs'

const row = { display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, fontSize: 12 } as const
const cellName = (c: { r: number; c: number }): string => `${colName(c.c)}${c.r + 1}`
/** "B2, D4:E5" -> the individual cells (max 32, like Excel). */
function parseCells(text: string): { r: number; c: number }[] | null {
  const ranges = []
  for (const part of text.split(',').map((s) => s.trim()).filter(Boolean)) {
    const p = parseRef(part)
    if (!p) return null
    ranges.push(p.range)
  }
  const cells = cellsOfRanges(ranges)
  return cells.length ? cells : null
}

export function DataTableDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const sel = sheet.selection.ranges[sheet.selection.ranges.length - 1]
  const [rowIn, setRowIn] = useState('')
  const [colIn, setColIn] = useState('')
  const [err, setErr] = useState('')
  const single = (t: string): { r: number; c: number } | null | undefined => {
    if (!t.trim()) return null
    const p = parseRef(t)
    return p && p.range.r1 === p.range.r2 && p.range.c1 === p.range.c2 ? { r: p.range.r1, c: p.range.c1 } : undefined
  }
  const ok = (): void => {
    const rIn = single(rowIn), cIn = single(colIn)
    if (rIn === undefined || cIn === undefined) return setErr('Input cell reference is not valid.')
    if (!rIn && !cIn) return setErr('Input cell reference is not valid.')
    const g = { r1: Math.min(sel.r1, sel.r2), c1: Math.min(sel.c1, sel.c2), r2: Math.max(sel.r1, sel.r2), c2: Math.max(sel.c1, sel.c2) }
    if (g.r2 - g.r1 < 1 || g.c2 - g.c1 < 1) return setErr('Select the whole table first: the input values and formulas plus the cells to fill.')
    onClose()
    computeDataTable(sheet, g, rIn, cIn)
  }
  return (
    <Dialog title="Data Table" width={320} onOk={ok} onCancel={onClose}>
      <div style={{ fontSize: 12 }}>
        {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
        <div style={row}><span style={{ width: 110 }}>Row input cell:</span><input autoFocus style={{ flex: 1 }} value={rowIn} onChange={(e) => { setRowIn(e.target.value); setErr('') }} /></div>
        <div style={row}><span style={{ width: 110 }}>Column input cell:</span><input style={{ flex: 1 }} value={colIn} onChange={(e) => { setColIn(e.target.value); setErr('') }} /></div>
        <div style={{ color: '#666' }}>Table range: {rangeA1(sel)}</div>
      </div>
    </Dialog>
  )
}

export function ScenarioManagerDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const [, tick] = useState(0)
  const list = scenariosOf(sheet)
  const [pick, setPick] = useState(list[0]?.name ?? '')
  const [edit, setEdit] = useState<null | { orig?: string; name: string; cells: string; values: string[]; comment: string }>(null)
  const [summary, setSummary] = useState<null | string>(null)
  const [err, setErr] = useState('')
  const cur = list.find((s) => s.name === pick)
  const refresh = (): void => tick((n) => n + 1)

  if (summary !== null) {
    return (
      <Dialog title="Scenario Summary" width={340} onOk={() => {
        const cells = summary.trim() ? parseCells(summary) : []
        if (!cells) return setErr('Reference is not valid.')
        onClose()
        scenarioSummary(sheet, cells)
      }} onCancel={() => { setSummary(null); setErr('') }}>
        <div style={{ fontSize: 12 }}>
          {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
          <div style={row}><span style={{ width: 90 }}>Result cells:</span><input autoFocus style={{ flex: 1 }} value={summary} onChange={(e) => { setSummary(e.target.value); setErr('') }} placeholder="B7, B8" /></div>
        </div>
      </Dialog>
    )
  }

  if (edit) {
    const cells = parseCells(edit.cells)
    const save = (): void => {
      if (!edit.name.trim()) return setErr('Enter a name for the scenario.')
      if (!cells) return setErr('Enter the changing cells, for example B1:B3.')
      if (list.some((s) => s.name === edit.name.trim() && s.name !== edit.orig)) return setErr('A scenario with this name already exists.')
      saveScenario(sheet, { name: edit.name.trim(), cells, values: cells.map((_c, i) => edit.values[i] ?? ''), comment: edit.comment || undefined }, edit.orig)
      setPick(edit.name.trim()); setEdit(null); setErr('')
    }
    return (
      <Dialog title={edit.orig ? 'Edit Scenario' : 'Add Scenario'} width={380} onOk={save} onCancel={() => { setEdit(null); setErr('') }}>
        <div style={{ fontSize: 12 }}>
          {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
          <div style={row}><span style={{ width: 110 }}>Scenario name:</span><input autoFocus style={{ flex: 1 }} value={edit.name} onChange={(e) => { setEdit({ ...edit, name: e.target.value }); setErr('') }} /></div>
          <div style={row}><span style={{ width: 110 }}>Changing cells:</span><input style={{ flex: 1 }} value={edit.cells} onChange={(e) => {
            const cs = parseCells(e.target.value)
            setEdit({ ...edit, cells: e.target.value, values: cs ? cs.map((c, i) => edit.values[i] ?? store.getRaw(sheet, c.r, c.c)) : edit.values }); setErr('')
          }} /></div>
          <div style={row}><span style={{ width: 110 }}>Comment:</span><input style={{ flex: 1 }} value={edit.comment} onChange={(e) => setEdit({ ...edit, comment: e.target.value })} /></div>
          {cells && (
            <div style={{ border: '1px solid #c8c8c8', maxHeight: 130, overflow: 'auto', padding: 4 }}>
              <div style={{ marginBottom: 4 }}>Scenario Values — type a value for each changing cell:</div>
              {cells.map((c, i) => (
                <div key={i} style={row}><span style={{ width: 50 }}>{cellName(c)}:</span><input style={{ flex: 1 }} value={edit.values[i] ?? ''} onChange={(e) => { const v = [...edit.values]; v[i] = e.target.value; setEdit({ ...edit, values: v }) }} /></div>
              ))}
            </div>
          )}
        </div>
      </Dialog>
    )
  }

  const startAdd = (): void => {
    const sel = sheet.selection.ranges[sheet.selection.ranges.length - 1]
    const text = rangeA1(sel)
    const cs = parseCells(text) ?? []
    setEdit({ name: '', cells: text, values: cs.map((c) => store.getRaw(sheet, c.r, c.c)), comment: '' })
  }
  const startEdit = (s: ScenarioDef): void => setEdit({ orig: s.name, name: s.name, cells: s.cells.map(cellName).join(', '), values: [...s.values], comment: s.comment ?? '' })
  const btn = (label: string, on: () => void, disabled = false): React.ReactElement => (
    <button className="xl-btn" style={{ display: 'block', width: '100%', marginBottom: 6 }} disabled={disabled} onClick={on}>{label}</button>
  )
  return (
    <Dialog title="Scenario Manager" width={430} onCancel={onClose} cancelLabel="Close">
      <div style={{ display: 'flex', gap: 12, fontSize: 12 }}>
        <div style={{ flex: 1 }}>
          <div>Scenarios:</div>
          <select size={8} style={{ width: '100%' }} value={pick} onChange={(e) => setPick(e.target.value)} onDoubleClick={() => cur && showScenario(sheet, cur)}>
            {list.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}
          </select>
          <div style={{ marginTop: 6 }}>Changing cells: {cur ? cur.cells.map(cellName).join(', ') : ''}</div>
          <div>Comment: {cur?.comment ?? ''}</div>
          {!list.length && <div style={{ color: '#666', marginTop: 6 }}>No Scenarios defined. Choose Add to add scenarios.</div>}
        </div>
        <div style={{ width: 110 }}>
          {btn('Add...', startAdd)}
          {btn('Delete', () => { if (cur) { deleteScenario(sheet, cur.name); setPick(list.find((s) => s.name !== cur.name)?.name ?? ''); refresh() } }, !cur)}
          {btn('Edit...', () => cur && startEdit(cur), !cur)}
          {btn('Summary...', () => setSummary(''), !list.length)}
          <div style={{ height: 8 }} />
          {btn('Show', () => cur && showScenario(sheet, cur), !cur)}
        </div>
      </div>
    </Dialog>
  )
}

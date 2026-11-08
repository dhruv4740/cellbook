// Data > Group / Ungroup / Show Detail / Hide Detail (Rows or Columns) and the Outline Settings dialog.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { applyOutline, OutlineMode, setOutlineSummary } from '../grid/outlineCommands'

const TITLES: Record<OutlineMode, string> = { group: 'Group', ungroup: 'Ungroup', show: 'Show Detail', hide: 'Hide Detail' }

export function OutlineAxisDialog({ mode, onClose }: { mode: OutlineMode; onClose: () => void }) {
  const [axis, setAxis] = useState<'row' | 'col'>('row')
  return (
    <Dialog title={TITLES[mode]} width={250} onCancel={onClose} onOk={() => { onClose(); applyOutline(mode, axis) }}>
      <div style={{ fontSize: 12 }}>
        <label style={{ display: 'flex', gap: 6 }}><input type="radio" autoFocus checked={axis === 'row'} onChange={() => setAxis('row')} /> Rows</label>
        <label style={{ display: 'flex', gap: 6 }}><input type="radio" checked={axis === 'col'} onChange={() => setAxis('col')} /> Columns</label>
      </div>
    </Dialog>
  )
}

export function OutlineSettingsDialog({ onClose }: { onClose: () => void }) {
  const o = store.active.layout.outline
  const [below, setBelow] = useState(o?.summaryBelow !== false)
  const [right, setRight] = useState(o?.summaryRight !== false)
  return (
    <Dialog title="Settings" width={300} onCancel={onClose} onOk={() => { onClose(); setOutlineSummary(below, right) }}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 4 }}>Direction</div>
        <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={below} onChange={(e) => setBelow(e.target.checked)} /> Summary rows below detail</label>
        <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={right} onChange={(e) => setRight(e.target.checked)} /> Summary columns to right of detail</label>
      </div>
    </Dialog>
  )
}

// Bottom status bar: mode indicator, accessibility, selection stats, view buttons, zoom slider.
import { useState } from 'react'
import { useStore } from '../model/store'
import { normRange } from '../model/types'
import { useUi, setUi } from './ui'
import { Menu, MenuItem, useMenu } from '../components/Menu'
import { Dialog } from '../components/Dialog'
import { CheckmarkCircle20Regular, TableSimple20Regular, DocumentPageBreak20Regular, PageFit20Regular, Subtract20Regular, Add20Regular } from '@fluentui/react-icons'
import './StatusBar.css'

function fmtNum(n: number): string {
  return n.toLocaleString('en-US', { maximumFractionDigits: 2 })
}

export default function StatusBar() {
  const store = useStore()
  const ui = useUi()
  const sheet = store.active
  const [menu, openMenu, closeMenu] = useMenu()
  const [zoomInput, setZoomInput] = useState(String(sheet.layout.zoom))
  const [items, setItems] = useState({ average: true, count: true, sum: true, numericalCount: false, min: false, max: false })

  let sum = 0, count = 0, numCount = 0, min = Infinity, max = -Infinity
  // Clamp to the used range: a whole-row/column/sheet selection (Ctrl+A, header click) would
  // otherwise iterate up to MAX_ROWS*MAX_COLS cells here and hang the tab.
  const used = store.usedRange(sheet)
  for (const raw of used ? sheet.selection.ranges : []) {
    const rg = normRange(raw) // ranges dragged up/left arrive reversed
    for (let r = rg.r1; r <= Math.min(rg.r2, used!.r2); r++) {
      if (sheet.layout.hiddenRows?.[r]) continue // Excel's status bar only counts visible cells
      for (let c = rg.c1; c <= Math.min(rg.c2, used!.c2); c++) {
        if (sheet.layout.hiddenCols?.[c]) continue
        const v = store.getValue(sheet, r, c)
        if (v === null || v === '') continue
        count++
        if (typeof v === 'number') { sum += v; numCount++; if (v < min) min = v; if (v > max) max = v }
      }
    }
  }
  const showStats = count >= 2 && numCount >= 1

  const af = sheet.layout.autoFilter
  const recordsText = af ? `${af.range.r2 - af.range.r1 - af.filteredOutRows.length} of ${af.range.r2 - af.range.r1} records found` : null

  const zoom = sheet.layout.zoom
  const setZoom = (z: number) => store.updateLayout(sheet, 'Zoom', (l) => (l.zoom = Math.max(10, Math.min(400, Math.round(z)))))

  const toggle: MenuItem[] = [
    { label: 'Formula Cell Reference', checked: false, onClick: () => {} },
    'separator',
    { label: 'Accessibility: Investigate', checked: true, onClick: () => {} },
    'separator',
    { label: 'Signatures', checked: false, onClick: () => {} },
    { label: 'Information Management Policy', checked: false, onClick: () => {} },
    { label: 'Permissions', checked: false, onClick: () => {} },
    'separator',
    { label: 'Caps Lock', checked: false, onClick: () => {} },
    { label: 'Num Lock', checked: true, onClick: () => {} },
    { label: 'Scroll Lock', checked: ui.scrollLock, onClick: () => setUi({ scrollLock: !ui.scrollLock }) },
    { label: 'Fixed Decimal', checked: false, onClick: () => {} },
    'separator',
    { label: 'Overtype Mode', checked: false, onClick: () => {} },
    { label: 'End Mode', checked: ui.endMode, onClick: () => setUi({ endMode: !ui.endMode }) },
    'separator',
    { label: 'Cell Mode', checked: true, onClick: () => {} },
    'separator',
    { label: 'Auto Calculate', checked: true, onClick: () => {} },
    { label: 'Page Number', checked: false, onClick: () => {} },
    'separator',
    { label: 'Average', checked: items.average, onClick: () => setItems((s) => ({ ...s, average: !s.average })) },
    { label: 'Count', checked: items.count, onClick: () => setItems((s) => ({ ...s, count: !s.count })) },
    { label: 'Numerical Count', checked: items.numericalCount, onClick: () => setItems((s) => ({ ...s, numericalCount: !s.numericalCount })) },
    { label: 'Minimum', checked: items.min, onClick: () => setItems((s) => ({ ...s, min: !s.min })) },
    { label: 'Maximum', checked: items.max, onClick: () => setItems((s) => ({ ...s, max: !s.max })) },
    { label: 'Sum', checked: items.sum, onClick: () => setItems((s) => ({ ...s, sum: !s.sum })) },
    'separator',
    { label: 'Upload Status', checked: false, onClick: () => {} },
    { label: 'View Shortcuts', checked: true, onClick: () => {} },
    { label: 'Zoom Slider', checked: true, onClick: () => {} },
    { label: 'Zoom', checked: true, onClick: () => {} }
  ]

  return (
    <div className="xl-statusbar" onContextMenu={(e) => { e.preventDefault(); openMenu(e) }}>
      <div className="xl-status-left">
        <span className="xl-status-mode">
          {ui.selectMode === 'extend' ? 'Extend Selection' : ui.selectMode === 'add' ? 'Add to Selection' : ui.busy ? ui.busy : store.calcPending && store.mode === 'Ready' ? 'Calculate' : store.mode}
        </span>
        {ui.scrollLock && <span className="xl-status-records">Scroll Lock</span>}
        {ui.endMode && <span className="xl-status-records">End Mode</span>}
        {recordsText && <span className="xl-status-records">{recordsText}</span>}
        <span className="xl-status-a11y"><CheckmarkCircle20Regular fontSize={14} /> Accessibility: Good to go</span>
        {showStats && (
          <span className="xl-status-stats">
            {items.average && <span>Average: {fmtNum(sum / numCount)}</span>}
            {items.numericalCount && <span>Numerical Count: {numCount}</span>}
            {items.count && <span>Count: {count}</span>}
            {items.min && min !== Infinity && <span>Minimum: {fmtNum(min)}</span>}
            {items.max && max !== -Infinity && <span>Maximum: {fmtNum(max)}</span>}
            {items.sum && <span>Sum: {fmtNum(sum)}</span>}
          </span>
        )}
      </div>
      <div className="xl-status-right">
        <div className="xl-view-btns">
          <button className={'xl-viewbtn' + (ui.viewMode === 'normal' ? ' active' : '')} title="Normal" onClick={() => setUi({ viewMode: 'normal' })}><TableSimple20Regular /></button>
          <button className="xl-viewbtn" title="Page Layout"><PageFit20Regular /></button>
          <button className={'xl-viewbtn' + (ui.viewMode === 'pageBreak' ? ' active' : '')} title="Page Break Preview" onClick={() => setUi({ viewMode: 'pageBreak' })}><DocumentPageBreak20Regular /></button>
        </div>
        <div className="xl-zoom">
          <button className="xl-zoombtn" onClick={() => setZoom(zoom - 10)}><Subtract20Regular fontSize={12} /></button>
          <input className="xl-zoom-slider" type="range" min={10} max={400} value={zoom} onChange={(e) => setZoom(+e.target.value)} />
          <button className="xl-zoombtn" onClick={() => setZoom(zoom + 10)}><Add20Regular fontSize={12} /></button>
          <button className="xl-zoom-pct" onClick={() => setUi({ dialog: { kind: 'zoom' } })}>{zoom}%</button>
        </div>
      </div>
      {menu && <Menu x={menu.x} y={menu.y} items={toggle} onClose={closeMenu} />}
      {ui.dialog?.kind === 'zoom' && (
        <Dialog
          title="Zoom"
          width={260}
          onCancel={() => setUi({ dialog: null })}
          onOk={() => { const n = parseInt(zoomInput, 10); if (n) setZoom(n); setUi({ dialog: null }) }}
        >
          <div className="xl-zoom-dialog">
            {[200, 100, 75, 50, 25].map((z) => (
              <label key={z}><input type="radio" name="zoom" checked={+zoomInput === z} onChange={() => setZoomInput(String(z))} /> {z}%</label>
            ))}
            <label><input type="radio" name="zoom" checked={![200, 100, 75, 50, 25].includes(+zoomInput)} onChange={() => {}} /> Custom:
              <input className="xl-zoom-custom" value={zoomInput} onChange={(e) => setZoomInput(e.target.value.replace(/[^0-9]/g, ''))} />%
            </label>
          </div>
        </Dialog>
      )}
    </div>
  )
}

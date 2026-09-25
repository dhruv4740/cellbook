// Home > Format as Table / Table Design > Table Styles: the Light 1-21 / Medium 1-28 / Dark 1-11
// gallery, each swatch a small rendered thumbnail (header + 2 banded rows) built from the same
// TableStyleColors the grid itself paints with, so the swatch always matches what you'd actually get.
import { useEffect, useRef } from 'react'
import { createPortal } from 'react-dom'
import { LIGHT_STYLES, MEDIUM_STYLES, DARK_STYLES, TableStyleColors } from '../model/tableStyles'
import './TableStyleGallery.css'

function Swatch({ s, selected, onPick }: { s: TableStyleColors; selected: boolean; onPick: () => void }) {
  const band1 = s.band1 ?? '#ffffff'
  const band2 = s.band2 ?? '#ffffff'
  return (
    <button type="button" className={'xl-tsw' + (selected ? ' selected' : '')} title={s.name.replace('TableStyle', '').replace(/([A-Za-z])(\d)/, '$1 $2')} onClick={onPick}>
      <div className="xl-tsw-thumb" style={{ borderColor: s.border ?? '#d4d4d4' }}>
        <div className="xl-tsw-row" style={{ background: s.headerFill ?? '#ffffff', borderBottom: `1px solid ${s.border ?? '#d4d4d4'}` }} />
        <div className="xl-tsw-row" style={{ background: band1 }} />
        <div className="xl-tsw-row" style={{ background: band2 }} />
        <div className="xl-tsw-row" style={{ background: band1 }} />
      </div>
    </button>
  )
}

export function TableStyleGallery({ x, y, current, onPick, onClose }: { x: number; y: number; current?: string; onPick: (styleName: string) => void; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const onDown = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) onClose() }
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('mousedown', onDown)
    window.addEventListener('keydown', onKey)
    return () => { window.removeEventListener('mousedown', onDown); window.removeEventListener('keydown', onKey) }
  }, [onClose])

  const section = (label: string, styles: TableStyleColors[]) => (
    <div className="xl-tsg-section">
      <div className="xl-tsg-label">{label}</div>
      <div className="xl-tsg-grid">
        {styles.map((s) => <Swatch key={s.name} s={s} selected={s.name === current} onPick={() => onPick(s.name)} />)}
      </div>
    </div>
  )

  return createPortal(
    <div ref={ref} className="xl-tsg" style={{ left: x, top: y }}>
      {section('Light', LIGHT_STYLES)}
      {section('Medium', MEDIUM_STYLES)}
      {section('Dark', DARK_STYLES)}
      <div className="xl-tsg-footer" onClick={onClose}>New Table Style...</div>
    </div>,
    document.body
  )
}

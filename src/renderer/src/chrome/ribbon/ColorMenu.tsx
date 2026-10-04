// Excel's color picker popup: Theme Colors (10x6), Standard Colors, No Fill/Automatic, More Colors.
import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import { store } from '../../model/store'
import { OFFICE_COLORS } from '../../data/themes'

// Office 2023+ theme: 10 columns x 6 rows (base + 5 tint/shade steps).
function shade(hex: string, amt: number): string {
  const n = parseInt(hex.slice(1), 16)
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255
  const mix = (c: number) => (amt > 0 ? c + (255 - c) * amt : c * (1 + amt))
  r = Math.round(mix(r)); g = Math.round(mix(g)); b = Math.round(mix(b))
  return '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase()
}
/** Follows the workbook's chosen theme colours (Page Layout > Colors). */
const themeRows = (): string[][] => {
  const base = (store.theme.colors ?? OFFICE_COLORS).colors.map((c) => '#' + c)
  return [0, 0.8, 0.6, 0.4, -0.25, -0.5].map((amt, i) => (i === 0 ? base : base.map((c) => shade(c, amt))))
}
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0']

export function ColorMenu({
  x, y, onClose, onPick, mode
}: { x: number; y: number; onClose: () => void; onPick: (color: string | undefined) => void; mode: 'font' | 'fill' }) {
  const [more, setMore] = useState(false)
  const [custom, setCustom] = useState('#000000')
  useEffect(() => {
    const key = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.preventDefault(); onClose() } }
    window.addEventListener('keydown', key)
    return () => window.removeEventListener('keydown', key)
  }, [onClose])
  const Swatch = ({ c, title }: { c: string; title: string }) => (
    <div className="xl-color-swatch" title={title} style={{ background: c }} onMouseDown={(e) => { e.preventDefault(); onPick(c); onClose() }} />
  )
  return createPortal(
    <div className="xl-menu-layer" style={{ pointerEvents: 'auto' }} onMouseDown={onClose}>
      <div className="xl-color-menu" style={{ left: x, top: y }} onMouseDown={(e) => e.stopPropagation()}>
        {!more ? (
          <>
            <div
              className="xl-menu-item"
              onMouseDown={(e) => { e.preventDefault(); onPick(undefined); onClose() }}
            >
              <span className="xl-menu-label">{mode === 'font' ? 'Automatic' : 'No Fill'}</span>
            </div>
            <div className="xl-color-section">Theme Colors</div>
            <div className="xl-color-grid">
              {themeRows().map((row, ri) => row.map((c, ci) => <Swatch key={ri + '-' + ci} c={c} title="Theme Color" />))}
            </div>
            <div className="xl-color-section">Standard Colors</div>
            <div className="xl-color-grid" style={{ gridTemplateRows: '1fr' }}>
              {STANDARD.map((c) => <Swatch key={c} c={c} title={c} />)}
            </div>
            <div className="xl-menu-sep" />
            <div className="xl-menu-item" onMouseDown={(e) => { e.preventDefault(); setMore(true) }}>
              <span className="xl-menu-label">More Colors...</span>
            </div>
          </>
        ) : (
          <div className="xl-more-colors">
            <input type="color" value={custom} onChange={(e) => setCustom(e.target.value)} />
            <input className="xl-hex-input" value={custom} onChange={(e) => setCustom(e.target.value)} />
            <div className="xl-dialog-footer" style={{ padding: '6px 0 0' }}>
              <button className="xl-btn primary" onClick={() => { onPick(custom); onClose() }}>OK</button>
              <button className="xl-btn" onClick={onClose}>Cancel</button>
            </div>
          </div>
        )}
      </div>
    </div>,
    document.body
  )
}

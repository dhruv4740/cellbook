// Insert > Screenshot: pick another open window, or Screen Clipping (drag a region of the screen).
import { useEffect, useRef, useState } from 'react'
import { Dialog } from '../components/Dialog'
import { insertPictureFromDataUrl } from '../data/pictures'
import { setUi, showToast } from '../chrome/ui'

export function ScreenshotDialog({ onClose }: { onClose: () => void }) {
  const [wins, setWins] = useState<{ name: string; dataUrl: string }[] | null>(null)
  useEffect(() => { void (window.api?.screenWindows() ?? Promise.resolve([])).then(setWins).catch(() => setWins([])) }, [])
  const clip = async (): Promise<void> => {
    onClose()
    const src = await window.api?.screenGrab()
    if (!src) { showToast('Screen Clipping could not capture the screen.'); return }
    setUi({ dialog: { kind: 'screenClip', props: { src } } })
  }
  return (
    <Dialog title="Screenshot" width={560} onCancel={onClose} cancelLabel="Cancel" onOk={undefined}
      footerLeft={<button className="xl-btn" onClick={() => void clip()}>Screen Clipping</button>}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 6 }}>Available Windows</div>
        {wins === null && <div style={{ color: '#605e5c' }}>Looking for open windows…</div>}
        {wins && !wins.length && <div style={{ color: '#605e5c' }}>No other windows are open.</div>}
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, maxHeight: 320, overflowY: 'auto' }}>
          {wins?.map((w, i) => (
            <div key={i} data-screenshot-window={w.name} role="button" style={{ width: 160, cursor: 'pointer', border: '1px solid #d1d1d1', padding: 4 }}
              onClick={() => { onClose(); void insertPictureFromDataUrl(w.dataUrl, w.name) }}>
              <img src={w.dataUrl} style={{ width: '100%', height: 90, objectFit: 'contain', background: '#f3f2f1', display: 'block' }} />
              <div style={{ marginTop: 3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={w.name}>{w.name}</div>
            </div>
          ))}
        </div>
      </div>
    </Dialog>
  )
}

/** Full-window overlay showing the captured screen; dragging a rectangle inserts that region as a picture. Esc cancels. */
export function ScreenClipOverlay({ src, onClose }: { src: string; onClose: () => void }) {
  const [rect, setRect] = useState<{ x0: number; y0: number; x1: number; y1: number } | null>(null)
  const box = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const k = (e: KeyboardEvent): void => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [onClose])
  const at = (e: React.MouseEvent): { x: number; y: number } => { const b = box.current!.getBoundingClientRect(); return { x: e.clientX - b.left, y: e.clientY - b.top } }
  const finish = (): void => {
    if (!rect) return onClose()
    const b = box.current!.getBoundingClientRect()
    const x = Math.min(rect.x0, rect.x1), y = Math.min(rect.y0, rect.y1), w = Math.abs(rect.x1 - rect.x0), h = Math.abs(rect.y1 - rect.y0)
    onClose()
    if (w < 4 || h < 4) return
    const img = new Image()
    img.onload = () => {
      const kx = img.naturalWidth / b.width, ky = img.naturalHeight / b.height
      const cv = document.createElement('canvas')
      cv.width = Math.round(w * kx); cv.height = Math.round(h * ky)
      cv.getContext('2d')!.drawImage(img, x * kx, y * ky, w * kx, h * ky, 0, 0, cv.width, cv.height)
      void insertPictureFromDataUrl(cv.toDataURL('image/png'), 'Screen Clipping')
    }
    img.src = src
  }
  const sel = rect ? { left: Math.min(rect.x0, rect.x1), top: Math.min(rect.y0, rect.y1), width: Math.abs(rect.x1 - rect.x0), height: Math.abs(rect.y1 - rect.y0) } : null
  return (
    <div ref={box} data-screen-clip style={{ position: 'fixed', inset: 0, zIndex: 100000, cursor: 'crosshair', background: `url(${src}) center / 100% 100% no-repeat`, userSelect: 'none' }}
      onMouseDown={(e) => { const p = at(e); setRect({ x0: p.x, y0: p.y, x1: p.x, y1: p.y }) }}
      onMouseMove={(e) => { if (rect && e.buttons) { const p = at(e); setRect({ ...rect, x1: p.x, y1: p.y }) } }}
      onMouseUp={finish}>
      <div style={{ position: 'absolute', inset: 0, background: 'rgba(255,255,255,0.55)', pointerEvents: 'none', clipPath: sel ? `polygon(0 0, 100% 0, 100% 100%, 0 100%, 0 0, ${sel.left}px ${sel.top}px, ${sel.left}px ${sel.top + sel.height}px, ${sel.left + sel.width}px ${sel.top + sel.height}px, ${sel.left + sel.width}px ${sel.top}px, ${sel.left}px ${sel.top}px)` : undefined }} />
      {sel && <div style={{ position: 'absolute', ...sel, border: '1px solid #217346', pointerEvents: 'none' }} />}
    </div>
  )
}

// Formulas > Watch Window: a small floating list of cells whose values you want to keep an eye on while working elsewhere.
import { useRef, useState } from 'react'
import { store, useStore } from '../model/store'
import { colName } from '../model/types'
import { setUi, useUi } from './ui'
import { parseRef } from '../dialogs/DataToolDialogs'

export default function WatchWindow() {
  useStore()
  const ui = useUi()
  const [pos, setPos] = useState({ x: 80, y: 260 })
  const [sel, setSel] = useState<number | null>(null)
  const [adding, setAdding] = useState<string | null>(null)
  const drag = useRef<{ mx: number; my: number; x: number; y: number } | null>(null)

  const onDown = (e: React.MouseEvent): void => {
    drag.current = { mx: e.clientX, my: e.clientY, x: pos.x, y: pos.y }
    const move = (ev: MouseEvent): void => { const d = drag.current; if (d) setPos({ x: d.x + ev.clientX - d.mx, y: d.y + ev.clientY - d.my }) }
    const up = (): void => { drag.current = null; window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up) }
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up)
  }
  const addWatch = (text: string): void => {
    const p = parseRef(text)
    if (!p) return
    const items = [...ui.watchItems]
    for (let r = p.range.r1; r <= p.range.r2 && items.length < 500; r++)
      for (let c = p.range.c1; c <= p.range.c2 && items.length < 500; c++)
        if (!items.some((w) => w.sheetUid === p.sheet.uid && w.r === r && w.c === c)) items.push({ sheetUid: p.sheet.uid, r, c })
    setUi({ watchItems: items })
    setAdding(null)
  }
  const th = { textAlign: 'left', padding: '2px 6px', borderBottom: '1px solid #d0d0d0', background: '#f3f3f3', fontWeight: 400, whiteSpace: 'nowrap' } as const
  const td = { padding: '2px 6px', whiteSpace: 'nowrap', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis' } as const
  return (
    <div style={{ position: 'fixed', left: pos.x, top: pos.y, width: 560, background: '#fff', border: '1px solid #8a8a8a', boxShadow: '0 4px 16px rgba(0,0,0,0.25)', zIndex: 900, fontSize: 12 }}>
      <div onMouseDown={onDown} style={{ display: 'flex', alignItems: 'center', background: '#f3f3f3', padding: '4px 8px', cursor: 'move', userSelect: 'none' }}>
        <span style={{ flex: 1 }}>Watch Window</span>
        <span role="button" aria-label="Close Watch Window" style={{ cursor: 'pointer', padding: '0 4px' }} onMouseDown={(e) => e.stopPropagation()} onClick={() => setUi({ watchOpen: false })}>✕</span>
      </div>
      <div style={{ padding: 6, display: 'flex', gap: 6, alignItems: 'center' }}>
        {adding === null ? (
          <button className="xl-btn" onClick={() => { const a = store.active.selection.active; setAdding(`${colName(a.c)}${a.r + 1}`) }}>Add Watch...</button>
        ) : (
          <>
            <input autoFocus value={adding} onChange={(e) => setAdding(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') addWatch(adding); if (e.key === 'Escape') setAdding(null) }} style={{ width: 160 }} />
            <button className="xl-btn" onClick={() => addWatch(adding)}>Add</button>
            <button className="xl-btn" onClick={() => setAdding(null)}>Cancel</button>
          </>
        )}
        <button className="xl-btn" disabled={sel === null} onClick={() => { if (sel !== null) { setUi({ watchItems: ui.watchItems.filter((_w, i) => i !== sel) }); setSel(null) } }}>Delete Watch</button>
      </div>
      <div style={{ maxHeight: 220, overflow: 'auto', borderTop: '1px solid #d0d0d0' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead><tr><th style={th}>Book</th><th style={th}>Sheet</th><th style={th}>Cell</th><th style={th}>Value</th><th style={th}>Formula</th></tr></thead>
          <tbody>
            {ui.watchItems.map((w, i) => {
              const sh = store.sheetByUid(w.sheetUid)
              if (!sh) return null
              const raw = store.getRaw(sh, w.r, w.c)
              return (
                <tr key={i} onClick={() => setSel(i)} onDoubleClick={() => { store.setActiveSheet(store.sheets.indexOf(sh)); store.setSelection({ ranges: [{ r1: w.r, c1: w.c, r2: w.r, c2: w.c }], active: { r: w.r, c: w.c } }, sh) }} style={{ background: sel === i ? '#cfe3d6' : undefined, cursor: 'default' }}>
                  <td style={td}>{store.fileName || 'Book1'}</td><td style={td}>{sh.name}</td><td style={td}>{colName(w.c)}{w.r + 1}</td>
                  <td style={td}>{store.getDisplay(sh, w.r, w.c).text}</td><td style={td}>{raw.startsWith('=') ? raw : ''}</td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

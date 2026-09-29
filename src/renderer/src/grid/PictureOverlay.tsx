// Floats pictures over the grid: positioned from a cell anchor (scroll / zoom / frozen-pane aware), with Excel's
// selection frame, drag-to-move, resize handles (corners keep the aspect ratio), Delete, and a right-click menu.
import { useEffect, useRef, useState } from 'react'
import { useStore, SheetState } from '../model/store'
import { setUi, useUi } from '../chrome/ui'
import { PictureDef } from '../model/types'
import type { ChartGeometry } from '../charts/ChartOverlay'
import { Menu } from '../components/Menu'
import { picturesOf, updatePicture, deletePicture, reorderPicture } from '../data/pictures'

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
type Handle = (typeof HANDLES)[number]
type Box = { left: number; top: number; width: number; height: number }

export function PictureOverlay({ sheet, geo }: { sheet: SheetState; geo: ChartGeometry }) {
  useStore()
  const ui = useUi()
  const pics = picturesOf(sheet)
  const selectedId = ui.selectedPicture
  const [menu, setMenu] = useState<{ x: number; y: number; id: string } | null>(null)
  const { colsAxis, rowsAxis, frozen, scroll, z, rowHeaderW, headerH, viewport } = geo
  const frozenW = colsAxis.totalSize(frozen.cols) * z
  const frozenH = rowsAxis.totalSize(frozen.rows) * z

  useEffect(() => {
    if (!selectedId) return
    const onKey = (e: KeyboardEvent): void => {
      if (e.key !== 'Delete' && e.key !== 'Backspace') return
      const el = document.activeElement
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA')) return
      deletePicture(sheet, selectedId)
      setUi({ selectedPicture: null })
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [selectedId, sheet])

  if (!pics.length) return null
  const sx = (c: number): number => (c < frozen.cols ? colsAxis.offset(c) * z : frozenW + colsAxis.offset(c) * z - colsAxis.offset(frozen.cols) * z - scroll.x)
  const sy = (r: number): number => (r < frozen.rows ? rowsAxis.offset(r) * z : frozenH + rowsAxis.offset(r) * z - rowsAxis.offset(frozen.rows) * z - scroll.y)
  const lx = (v: number): number => (v < frozenW ? v / z : colsAxis.offset(frozen.cols) + (v - frozenW + scroll.x) / z)
  const ly = (v: number): number => (v < frozenH ? v / z : rowsAxis.offset(frozen.rows) + (v - frozenH + scroll.y) / z)
  const boxOf = (p: PictureDef): Box => ({ left: sx(p.anchor.c) + p.anchor.dx * z, top: sy(p.anchor.r) + p.anchor.dy * z, width: p.width * z, height: p.height * z })
  const commit = (p: PictureDef, b: Box): void => {
    const x = lx(b.left), y = ly(b.top)
    const c = colsAxis.indexAt(x), r = rowsAxis.indexAt(y)
    updatePicture(sheet, p.id, { anchor: { r, c, dx: x - colsAxis.offset(c), dy: y - rowsAxis.offset(r) }, width: Math.max(8, b.width / z), height: Math.max(8, b.height / z) }, 'Move or Resize Picture')
  }

  return (
    <div style={{ position: 'absolute', left: rowHeaderW, top: headerH, width: Math.max(0, viewport.w - rowHeaderW), height: Math.max(0, viewport.h - headerH), overflow: 'hidden', pointerEvents: 'none' }}>
      {pics.map((p) => (
        <PictureBox key={p.id} pic={p} box={boxOf(p)} selected={p.id === selectedId} onSelect={() => setUi({ selectedPicture: p.id, selectedChart: null })}
          onContext={(x, y) => { setUi({ selectedPicture: p.id, selectedChart: null }); setMenu({ x, y, id: p.id }) }} onCommit={(b) => commit(p, b)} />
      ))}
      {menu && (
        <div style={{ pointerEvents: 'auto' }}>
          <Menu x={menu.x} y={menu.y} onClose={() => setMenu(null)} items={[
            { label: 'Bring to &Front', onClick: () => reorderPicture(sheet, menu.id, true) },
            { label: 'S&end to Back', onClick: () => reorderPicture(sheet, menu.id, false) },
            'separator',
            { label: '&Delete', onClick: () => { deletePicture(sheet, menu.id); setUi({ selectedPicture: null }) } }
          ]} />
        </div>
      )}
    </div>
  )
}

function PictureBox({ pic, box, selected, onSelect, onContext, onCommit }: { pic: PictureDef; box: Box; selected: boolean; onSelect: () => void; onContext: (x: number, y: number) => void; onCommit: (b: Box) => void }) {
  const [drag, setDrag] = useState<Box | null>(null)
  const start = useRef<{ mx: number; my: number; box: Box; handle?: Handle } | null>(null)
  const live = drag ?? box

  function begin(e: React.MouseEvent, handle?: Handle): void {
    e.preventDefault()
    e.stopPropagation()
    ;(e.currentTarget as HTMLElement).closest<HTMLElement>('.grid-host')?.focus() // keep Ctrl+Z / Delete working
    onSelect()
    start.current = { mx: e.clientX, my: e.clientY, box, handle }
    const move = (ev: MouseEvent): void => {
      const s = start.current
      if (!s) return
      const dx = ev.clientX - s.mx, dy = ev.clientY - s.my
      const b = { ...s.box }
      if (!s.handle) {
        b.left = Math.max(0, s.box.left + dx)
        b.top = Math.max(0, s.box.top + dy)
      } else {
        const ratio = s.box.width / s.box.height
        if (s.handle.includes('e')) b.width = Math.max(16, s.box.width + dx)
        if (s.handle.includes('w')) { b.width = Math.max(16, s.box.width - dx); b.left = s.box.left + s.box.width - b.width }
        if (s.handle.includes('s')) b.height = Math.max(16, s.box.height + dy)
        if (s.handle.includes('n')) { b.height = Math.max(16, s.box.height - dy); b.top = s.box.top + s.box.height - b.height }
        // Corner handles keep the picture's proportions; Shift frees them.
        if (s.handle.length === 2 && !ev.shiftKey) {
          b.height = b.width / ratio
          if (s.handle.includes('n')) b.top = s.box.top + s.box.height - b.height
        }
      }
      setDrag(b)
    }
    const up = (): void => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      start.current = null
      setDrag((cur) => { if (cur) onCommit(cur); return null })
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }

  return (
    <div
      style={{ position: 'absolute', left: live.left, top: live.top, width: live.width, height: live.height, pointerEvents: 'auto', outline: selected ? '1px solid #217346' : 'none', cursor: drag ? 'move' : 'default' }}
      onMouseDown={(e) => begin(e)}
      onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); onContext(e.clientX, e.clientY) }}
    >
      <img src={pic.src} alt={pic.alt ?? pic.name} draggable={false} style={{ width: '100%', height: '100%', display: 'block', userSelect: 'none' }} />
      {selected && HANDLES.map((h) => (
        <div key={h} onMouseDown={(e) => begin(e, h)} style={{ position: 'absolute', width: 7, height: 7, background: '#fff', border: '1px solid #217346', borderRadius: '50%', ...handlePos(h) }} />
      ))}
    </div>
  )
}

function handlePos(h: Handle): React.CSSProperties {
  const mid = { transform: 'translate(-50%, -50%)' } as const
  switch (h) {
    case 'nw': return { left: 0, top: 0, ...mid, cursor: 'nwse-resize' }
    case 'n': return { left: '50%', top: 0, ...mid, cursor: 'ns-resize' }
    case 'ne': return { left: '100%', top: 0, ...mid, cursor: 'nesw-resize' }
    case 'e': return { left: '100%', top: '50%', ...mid, cursor: 'ew-resize' }
    case 'se': return { left: '100%', top: '100%', ...mid, cursor: 'nwse-resize' }
    case 's': return { left: '50%', top: '100%', ...mid, cursor: 'ns-resize' }
    case 'sw': return { left: 0, top: '100%', ...mid, cursor: 'nesw-resize' }
    case 'w': return { left: 0, top: '50%', ...mid, cursor: 'ew-resize' }
  }
}

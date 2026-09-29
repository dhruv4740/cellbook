// A floating object frame (picture, shape, text box): Excel's selection outline, eight resize handles, drag-to-move.
// Corner handles keep the aspect ratio when `lockAspect` is set (Shift toggles it), like Excel.
import { ReactNode, useRef, useState } from 'react'

export type Box = { left: number; top: number; width: number; height: number }
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
type Handle = (typeof HANDLES)[number]

export function FloatBox({ box, selected, lockAspect, minSize = 16, onSelect, onContext, onCommit, onDoubleClick, children }: {
  box: Box; selected: boolean; lockAspect?: boolean; minSize?: number
  onSelect: () => void; onContext: (x: number, y: number) => void; onCommit: (b: Box) => void; onDoubleClick?: () => void
  children: ReactNode
}) {
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
        if (s.handle.includes('e')) b.width = Math.max(minSize, s.box.width + dx)
        if (s.handle.includes('w')) { b.width = Math.max(minSize, s.box.width - dx); b.left = s.box.left + s.box.width - b.width }
        if (s.handle.includes('s')) b.height = Math.max(minSize, s.box.height + dy)
        if (s.handle.includes('n')) { b.height = Math.max(minSize, s.box.height - dy); b.top = s.box.top + s.box.height - b.height }
        if (s.handle.length === 2 && !!lockAspect !== ev.shiftKey) {
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
      onDoubleClick={onDoubleClick}
      onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); onContext(e.clientX, e.clientY) }}
    >
      {children}
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

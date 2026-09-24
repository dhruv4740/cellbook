// Excel-style popup menu (right-click menus, dropdowns). Shared by grid, sheet tabs, ribbon and status bar.
import { ReactNode, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import './Menu.css'

export type MenuItem =
  | 'separator'
  | {
      label: string // Excel's exact wording; '&' marks the access-key letter, e.g. 'Cu&t'
      icon?: ReactNode
      shortcut?: string
      disabled?: boolean
      checked?: boolean
      submenu?: MenuItem[]
      /** Custom content rendered instead of a row (e.g. the 'Paste Options:' icon strip). */
      render?: (close: () => void) => ReactNode
      onClick?: () => void
    }

export interface MenuProps {
  x: number
  y: number
  items: MenuItem[]
  onClose: () => void
  /** Optional header (e.g. the mini formatting toolbar above the cell context menu). */
  header?: ReactNode
}

function Label({ text }: { text: string }) {
  const i = text.indexOf('&')
  if (i < 0) return <>{text}</>
  return (
    <>
      {text.slice(0, i)}
      <u>{text[i + 1]}</u>
      {text.slice(i + 2)}
    </>
  )
}

function MenuList({ x, y, items, onClose, depth }: { x: number; y: number; items: MenuItem[]; onClose: () => void; depth: number }) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ x, y })
  const [open, setOpen] = useState<number | null>(null)
  const [subPos, setSubPos] = useState({ x: 0, y: 0 })

  // Keep the menu on screen.
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const r = el.getBoundingClientRect()
    setPos({
      x: x + r.width > window.innerWidth ? Math.max(0, depth ? x - r.width - 200 : window.innerWidth - r.width - 2) : x,
      y: y + r.height > window.innerHeight ? Math.max(0, window.innerHeight - r.height - 2) : y
    })
  }, [x, y, depth])

  return (
    <div ref={ref} className="xl-menu" style={{ left: pos.x, top: pos.y }} onMouseDown={(e) => e.stopPropagation()} onContextMenu={(e) => e.preventDefault()}>
      {items.map((it, i) => {
        if (it === 'separator') return <div key={i} className="xl-menu-sep" />
        if (it.render) return <div key={i}>{it.render(onClose)}</div>
        return (
          <div
            key={i}
            className={'xl-menu-item' + (it.disabled ? ' disabled' : '') + (open === i ? ' open' : '')}
            onMouseEnter={(e) => {
              if (it.submenu) {
                const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
                setSubPos({ x: r.right - 2, y: r.top - 4 })
                setOpen(i)
              } else setOpen(null)
            }}
            onClick={() => {
              if (it.disabled || it.submenu) return
              onClose()
              it.onClick?.()
            }}
          >
            <span className="xl-menu-icon">{it.checked ? '✓' : it.icon}</span>
            <span className="xl-menu-label"><Label text={it.label} /></span>
            <span className="xl-menu-shortcut">{it.shortcut}</span>
            <span className="xl-menu-arrow">{it.submenu ? '›' : ''}</span>
          </div>
        )
      })}
      {open !== null && items[open] !== 'separator' && (items[open] as { submenu?: MenuItem[] }).submenu && (
        <MenuList x={subPos.x} y={subPos.y} items={(items[open] as { submenu: MenuItem[] }).submenu} onClose={onClose} depth={depth + 1} />
      )}
    </div>
  )
}

export function Menu({ x, y, items, onClose, header }: MenuProps) {
  useEffect(() => {
    const close = () => onClose()
    const key = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('mousedown', close)
    window.addEventListener('blur', close)
    window.addEventListener('keydown', key)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('blur', close)
      window.removeEventListener('keydown', key)
    }
  }, [onClose])
  return createPortal(
    <div className="xl-menu-layer">
      {header && (
        <div className="xl-menu-header" style={{ left: x, top: y - 70 }} onMouseDown={(e) => e.stopPropagation()}>
          {header}
        </div>
      )}
      <MenuList x={x} y={y} items={items} onClose={onClose} depth={0} />
    </div>,
    document.body
  )
}

/** Hook for components that show a context menu: const [menu, openMenu, closeMenu] = useMenu() */
export function useMenu(): [{ x: number; y: number } | null, (e: { clientX: number; clientY: number }) => void, () => void] {
  const [pos, setPos] = useState<{ x: number; y: number } | null>(null)
  return [pos, (e) => setPos({ x: e.clientX, y: e.clientY }), () => setPos(null)]
}

// Excel-style modal dialog: grey title bar with '?' and close, content, right-aligned OK / Cancel.
import { ReactNode, useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import './Dialog.css'

export interface DialogProps {
  title: string
  children: ReactNode
  onOk?: () => void // omit to hide the OK button
  onCancel: () => void
  okLabel?: string
  cancelLabel?: string
  width?: number
  /** Extra buttons on the left of the footer. */
  footerLeft?: ReactNode
  hideFooter?: boolean
}

export function Dialog({ title, children, onOk, onCancel, okLabel = 'OK', cancelLabel = 'Cancel', width = 400, footerLeft, hideFooter }: DialogProps) {
  const [pos, setPos] = useState<{ x: number; y: number } | null>(null)
  const drag = useRef<{ dx: number; dy: number } | null>(null)
  const box = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel()
      else if (e.key === 'Enter' && onOk && !(e.target instanceof HTMLTextAreaElement)) onOk()
    }
    window.addEventListener('keydown', key)
    return () => window.removeEventListener('keydown', key)
  }, [onOk, onCancel])

  useEffect(() => {
    const move = (e: MouseEvent) => drag.current && setPos({ x: e.clientX - drag.current.dx, y: e.clientY - drag.current.dy })
    const up = () => (drag.current = null)
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
    return () => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
    }
  }, [])

  return createPortal(
    <div className="xl-dialog-overlay">
      <div ref={box} className="xl-dialog" style={{ width, ...(pos ? { left: pos.x, top: pos.y, transform: 'none' } : {}) }}>
        <div
          className="xl-dialog-title"
          onMouseDown={(e) => {
            const r = box.current!.getBoundingClientRect()
            drag.current = { dx: e.clientX - r.left, dy: e.clientY - r.top }
          }}
        >
          <span>{title}</span>
          <span className="xl-dialog-tbtn">?</span>
          <span className="xl-dialog-tbtn" onClick={onCancel}>✕</span>
        </div>
        <div className="xl-dialog-body">{children}</div>
        {!hideFooter && (
          <div className="xl-dialog-footer">
            <div style={{ flex: 1 }}>{footerLeft}</div>
            {onOk && <button className="xl-btn primary" onClick={onOk}>{okLabel}</button>}
            <button className="xl-btn" onClick={onCancel}>{cancelLabel}</button>
          </div>
        )}
      </div>
    </div>,
    document.body
  )
}

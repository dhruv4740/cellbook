// Generic ribbon building blocks: big/small buttons, split buttons, combo boxes, galleries.
import { ReactNode, useState, useRef, useEffect, CSSProperties } from 'react'
import { Menu, MenuItem } from '../../components/Menu'
import { ChevronDown12Regular } from '@fluentui/react-icons'

export interface Tip {
  title: string
  desc?: string
  shortcut?: string
}

/** Excel ScreenTip: bold title (+ shortcut), then description, shown after a hover delay. */
export function useTip() {
  const [show, setShow] = useState(false)
  const t = useRef<number | undefined>(undefined)
  const onEnter = () => { t.current = window.setTimeout(() => setShow(true), 500) }
  const onLeave = () => { window.clearTimeout(t.current); setShow(false) }
  return { show, onEnter, onLeave }
}

export function ScreenTip({ tip }: { tip: Tip }) {
  return (
    <div className="xl-tip">
      <div className="xl-tip-title">{tip.title}{tip.shortcut ? <span className="xl-tip-shortcut">  {tip.shortcut}</span> : null}</div>
      {tip.desc && <div className="xl-tip-desc">{tip.desc}</div>}
    </div>
  )
}

function TipWrap({ tip, children }: { tip?: Tip; children: ReactNode }) {
  const { show, onEnter, onLeave } = useTip()
  return (
    <div className="xl-tipwrap" onMouseEnter={onEnter} onMouseLeave={onLeave} style={{ position: 'relative', display: 'contents' }}>
      {children}
      {show && tip && (
        <div style={{ position: 'absolute', top: '100%', left: 0, zIndex: 10000 }}>
          <ScreenTip tip={tip} />
        </div>
      )}
    </div>
  )
}

export interface BtnProps {
  icon: ReactNode
  label?: string
  tip?: Tip
  onClick?: () => void
  active?: boolean
  disabled?: boolean
  color?: string // small color swatch under the icon (Font Color / Fill Color)
}

/** Large ribbon button: icon on top, label wraps below, used for group headline commands. */
export function BigButton({ icon, label, tip, onClick, active, disabled, color, menu, wide }: BtnProps & { menu?: MenuItem[]; wide?: boolean }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLButtonElement>(null)
  return (
    <TipWrap tip={tip}>
      <button
        ref={ref}
        className={'xl-big-btn' + (wide ? ' wide' : '') + (active ? ' active' : '') + (open ? ' active' : '')}
        disabled={disabled}
        onClick={() => {
          if (menu) {
            const r = ref.current!.getBoundingClientRect()
            setOpen({ x: r.left, y: r.bottom })
          } else onClick?.()
        }}
      >
        <span className="xl-big-icon">
          {icon}
          {color && <span className="xl-color-bar" style={{ background: color }} />}
        </span>
        {label && <span className="xl-big-label">{label}{menu && <ChevronDown12Regular className="xl-drop-caret" />}</span>}
      </button>
      {open && menu && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
    </TipWrap>
  )
}

/** Split button: left = default action (icon+label), right = separate dropdown caret. */
export function SplitButton({ icon, label, tip, onClick, active, color, menu }: BtnProps & { menu: MenuItem[] }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLDivElement>(null)
  return (
    <TipWrap tip={tip}>
      <div ref={ref} className={'xl-big-btn xl-split' + (active ? ' active' : '')}>
        <button className="xl-split-main" onClick={onClick}>
          <span className="xl-big-icon">
            {icon}
            {color && <span className="xl-color-bar" style={{ background: color }} />}
          </span>
          {label && <span className="xl-big-label">{label}</span>}
        </button>
        <button
          className="xl-split-caret"
          onClick={() => {
            const r = ref.current!.getBoundingClientRect()
            setOpen({ x: r.left, y: r.bottom })
          }}
        >
          <ChevronDown12Regular />
        </button>
      </div>
      {open && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
    </TipWrap>
  )
}

/** Row split button: icon + label + caret in a single ~24px-tall row, used 3-stacked (Cells, Editing groups). */
export function StackSplit({ icon, label, tip, onClick, menu }: BtnProps & { menu: MenuItem[] }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLDivElement>(null)
  return (
    <TipWrap tip={tip}>
      <div ref={ref} className="xl-stacksplit">
        <button className="xl-stacksplit-main" onClick={onClick}>
          <span className="xl-stacksplit-icon">{icon}</span>
          {label && <span className="xl-stacksplit-label">{label}</span>}
        </button>
        <button
          className="xl-stacksplit-caret"
          onClick={() => {
            const r = ref.current!.getBoundingClientRect()
            setOpen({ x: r.left, y: r.bottom })
          }}
        >
          <ChevronDown12Regular />
        </button>
      </div>
      {open && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
    </TipWrap>
  )
}

/** Small (16px) ribbon button used stacked 2-3 per group, or in a row. */
export function SmallButton({ icon, label, tip, onClick, onDoubleClick, active, disabled, menu }: BtnProps & { menu?: MenuItem[]; onDoubleClick?: () => void }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLButtonElement>(null)
  return (
    <TipWrap tip={tip}>
      <button
        ref={ref}
        className={'xl-small-btn' + (active ? ' active' : '') + (open ? ' active' : '')}
        disabled={disabled}
        onDoubleClick={onDoubleClick}
        onClick={() => {
          if (menu) {
            const r = ref.current!.getBoundingClientRect()
            setOpen({ x: r.left, y: r.bottom })
          } else onClick?.()
        }}
      >
        <span className="xl-small-icon">{icon}</span>
        {label && <span className="xl-small-label">{label}</span>}
        {menu && <ChevronDown12Regular className="xl-drop-caret" />}
      </button>
      {open && menu && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
    </TipWrap>
  )
}

/** Group launcher arrow, bottom-right of a group, opens a dialog. */
export function DialogLauncher({ tip, onClick }: { tip: Tip; onClick: () => void }) {
  return (
    <TipWrap tip={tip}>
      <button className="xl-launcher" onClick={onClick}>⤡</button>
    </TipWrap>
  )
}

export function Group({ label, children, style }: { label: string; children: ReactNode; style?: CSSProperties }) {
  return (
    <div className="xl-group" style={style}>
      <div className="xl-group-body">{children}</div>
      <div className="xl-group-label">{label}</div>
    </div>
  )
}

export function GroupSep() {
  return <div className="xl-group-sep" />
}

/** Editable combo box (font name / size / number format). */
export function Combo({
  value, options, onChange, onPick, width = 110, tip, editable = true
}: {
  value: string
  options: string[]
  onChange?: (v: string) => void
  onPick?: (v: string) => void
  width?: number
  tip?: Tip
  editable?: boolean
}) {
  const [open, setOpen] = useState(false)
  const [text, setText] = useState(value)
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => setText(value), [value])
  useEffect(() => {
    if (!open) return
    const close = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) setOpen(false) }
    window.addEventListener('mousedown', close)
    return () => window.removeEventListener('mousedown', close)
  }, [open])
  return (
    <TipWrap tip={tip}>
      <div className="xl-combo" ref={ref} style={{ width }}>
        <input
          readOnly={!editable}
          value={text}
          onChange={(e) => setText(e.target.value)}
          onFocus={() => setOpen(true)}
          onClick={() => setOpen(true)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { onChange?.(text); (e.target as HTMLInputElement).blur(); setOpen(false) }
            if (e.key === 'Escape') { setText(value); setOpen(false) }
          }}
        />
        <button className="xl-combo-caret" onClick={() => setOpen((o) => !o)}><ChevronDown12Regular /></button>
        {open && (
          <div className="xl-combo-list">
            {options.map((o) => (
              <div
                key={o}
                className="xl-combo-item"
                style={/font/i.test(width + '') ? {} : undefined}
                onMouseDown={(e) => { e.preventDefault(); setText(o); onPick ? onPick(o) : onChange?.(o); setOpen(false) }}
              >
                {o}
              </div>
            ))}
          </div>
        )}
      </div>
    </TipWrap>
  )
}

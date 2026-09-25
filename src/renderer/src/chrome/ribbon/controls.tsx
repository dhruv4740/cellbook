// Generic ribbon building blocks: big/small buttons, split buttons, combo boxes, galleries.
import { ReactNode, useState, useRef, useEffect, useLayoutEffect, useContext, createContext, CSSProperties, RefObject } from 'react'
import { createPortal } from 'react-dom'
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
  /** Excel KeyTip access-key sequence for this command (e.g. 'FN', 'B', 'AT'), shown as a badge
   *  over the button while KeyTips are active (Alt / F10). See chrome/KeyTips.tsx. */
  keytip?: string
}

/** The little yellow-bordered badge Excel draws over a command while KeyTips are active.
 *  Always rendered (so the KeyTips engine can find it by data-keytip) but hidden by CSS
 *  unless an ancestor carries .xl-keytips-active. */
export function KeyTipBadge({ code }: { code?: string }) {
  if (!code) return null
  return <span className="xl-keytip-badge" data-keytip={code}>{code}</span>
}

/** Large ribbon button: icon on top, label wraps below, used for group headline commands. */
export function BigButton({ icon, label, tip, onClick, active, disabled, color, menu, wide, keytip }: BtnProps & { menu?: MenuItem[]; wide?: boolean }) {
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
        <span className="xl-big-icon" style={{ position: 'relative' }}>
          {icon}
          {color && <span className="xl-color-bar" style={{ background: color }} />}
          <KeyTipBadge code={keytip} />
        </span>
        {label && <span className="xl-big-label">{label}{menu && <ChevronDown12Regular className="xl-drop-caret" />}</span>}
      </button>
      {open && menu && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
    </TipWrap>
  )
}

/** Split button: left = default action (icon+label), right = separate dropdown caret. */
export function SplitButton({ icon, label, tip, onClick, active, color, menu, keytip, menuOnly }: BtnProps & { menu: MenuItem[]; menuOnly?: boolean }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLDivElement>(null)
  const openMenu = (): void => {
    const r = ref.current!.getBoundingClientRect()
    setOpen({ x: r.left, y: r.bottom })
  }
  return (
    <TipWrap tip={tip}>
      <div ref={ref} className={'xl-big-btn xl-split' + (active ? ' active' : '')}>
        <button className="xl-split-main" onClick={menuOnly ? openMenu : onClick}>
          <span className="xl-big-icon" style={{ position: 'relative' }}>
            {icon}
            {color && <span className="xl-color-bar" style={{ background: color }} />}
            <KeyTipBadge code={keytip} />
          </span>
          {label && <span className="xl-big-label">{label}</span>}
        </button>
        <button
          className="xl-split-caret"
          onClick={openMenu}
        >
          <ChevronDown12Regular />
        </button>
      </div>
      {open && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
    </TipWrap>
  )
}

/** Row split button: icon + label + caret in a single ~24px-tall row, used 3-stacked (Cells, Editing groups). */
export function StackSplit({ icon, label, tip, onClick, menu, keytip }: BtnProps & { menu: MenuItem[] }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLDivElement>(null)
  return (
    <TipWrap tip={tip}>
      <div ref={ref} className="xl-stacksplit">
        <button className="xl-stacksplit-main" onClick={onClick}>
          <span className="xl-stacksplit-icon" style={{ position: 'relative' }}>{icon}<KeyTipBadge code={keytip} /></span>
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
export function SmallButton({ icon, label, tip, onClick, onDoubleClick, active, disabled, menu, keytip, split }: BtnProps & { menu?: MenuItem[]; onDoubleClick?: () => void; split?: boolean }) {
  const [open, setOpen] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLButtonElement>(null)
  const wrap = useRef<HTMLDivElement>(null)
  if (split && menu) {
    // Excel's Underline / Borders: the button runs the current choice, the arrow opens the list.
    return (
      <TipWrap tip={tip}>
        <div ref={wrap} className={'xl-small-btn xl-split' + (active ? ' active' : '')}>
          <button className="xl-small-split-main" disabled={disabled} onClick={onClick}>
            <span className="xl-small-icon" style={{ position: 'relative' }}>{icon}<KeyTipBadge code={keytip} /></span>
          </button>
          <button className="xl-small-split-caret" disabled={disabled} onClick={() => { const r = wrap.current!.getBoundingClientRect(); setOpen({ x: r.left, y: r.bottom }) }}>
            <ChevronDown12Regular />
          </button>
        </div>
        {open && <Menu x={open.x} y={open.y} items={menu} onClose={() => setOpen(null)} />}
      </TipWrap>
    )
  }
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
        <span className="xl-small-icon" style={{ position: 'relative' }}>{icon}<KeyTipBadge code={keytip} /></span>
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

/* ---- Adaptive ribbon scaling ------------------------------------------------------------
 * Real Excel reduces groups right-to-left as the window narrows, eventually collapsing a
 * group into a single button (icon + name + chevron) that opens the group at full size in a
 * popup. RibbonScaleProvider measures each Group's natural width via an off-screen clone and,
 * given the available width, decides (right-to-left, since later-registered = further right)
 * which groups must collapse to fit. */
interface ScaleApi {
  collapsed: Set<string>
  report: (label: string, width: number) => void
}
const ScaleContext = createContext<ScaleApi | null>(null)

const GROUP_SEP_WIDTH = 3
const COLLAPSED_GROUP_WIDTH = 66

export function RibbonScaleProvider({ containerRef, children }: { containerRef: RefObject<HTMLDivElement | null>; children: ReactNode }) {
  const widthsRef = useRef<Map<string, number>>(new Map())
  const orderRef = useRef<string[]>([])
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set())
  const [, bump] = useState(0)

  const report = (label: string, width: number) => {
    if (!orderRef.current.includes(label)) orderRef.current.push(label)
    if (widthsRef.current.get(label) !== width) {
      widthsRef.current.set(label, width)
      bump((n) => n + 1)
    }
  }

  const recompute = () => {
    const el = containerRef.current
    if (!el) return
    const avail = el.clientWidth
    if (!avail) return
    const labels = orderRef.current
    let total = 0
    for (const l of labels) total += (widthsRef.current.get(l) ?? 0) + GROUP_SEP_WIDTH
    const next = new Set<string>()
    for (let i = labels.length - 1; i >= 0 && total > avail; i--) {
      const l = labels[i]
      const full = widthsRef.current.get(l) ?? 0
      if (full <= COLLAPSED_GROUP_WIDTH) continue
      total -= full - COLLAPSED_GROUP_WIDTH
      next.add(l)
    }
    setCollapsed((prev) => (prev.size === next.size && [...prev].every((l) => next.has(l)) ? prev : next))
  }

  useLayoutEffect(() => { recompute() })

  useEffect(() => {
    const el = containerRef.current
    if (!el || typeof ResizeObserver === 'undefined') return
    const ro = new ResizeObserver(() => recompute())
    ro.observe(el)
    return () => ro.disconnect()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return <ScaleContext.Provider value={{ collapsed, report }}>{children}</ScaleContext.Provider>
}

/** Portal panel anchored under a collapsed group's button, showing that group at full size. */
function GroupPopup({ anchor, onClose, children }: { anchor: HTMLElement | null; onClose: () => void; children: ReactNode }) {
  const [pos, setPos] = useState<{ x: number; y: number } | null>(null)
  const ref = useRef<HTMLDivElement>(null)

  useLayoutEffect(() => {
    if (!anchor) return
    const r = anchor.getBoundingClientRect()
    setPos({ x: r.left, y: r.bottom + 1 })
  }, [anchor])

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement
      if (ref.current?.contains(t) || anchor?.contains(t)) return
      onClose()
    }
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('mousedown', onDown)
    window.addEventListener('keydown', onKey)
    return () => { window.removeEventListener('mousedown', onDown); window.removeEventListener('keydown', onKey) }
  }, [anchor, onClose])

  // Clicking a leaf command (a plain button, not one that itself opens a submenu/caret) closes
  // the popup after the click is handled; clicking a dropdown caret, or a button that is itself
  // nothing but a dropdown trigger (e.g. Insert > Charts' gallery buttons), keeps it open so its
  // own menu can appear instead of being unmounted along with the popup.
  const onClickCapture = (e: { target: EventTarget | null }) => {
    const t = e.target as HTMLElement
    if (t.closest('[class*="caret"]')) return
    const btn = t.closest('button, .xl-combo-item') as HTMLElement | null
    if (!btn) return
    if (btn.querySelector('[class*="caret"]')) return
    onClose()
  }

  if (!pos) return null
  return createPortal(
    <div ref={ref} className="xl-group-popup" style={{ left: pos.x, top: pos.y }} onMouseDown={(e) => e.stopPropagation()} onClickCapture={onClickCapture}>
      {children}
    </div>,
    document.body
  )
}

export function Group({ label, children, style, icon }: { label: string; children: ReactNode; style?: CSSProperties; icon?: ReactNode }) {
  const scale = useContext(ScaleContext)
  const measureRef = useRef<HTMLDivElement>(null)
  const anchorRef = useRef<HTMLButtonElement>(null)
  const [open, setOpen] = useState(false)
  const collapsed = scale?.collapsed.has(label) ?? false

  useLayoutEffect(() => {
    if (!scale || !measureRef.current) return
    scale.report(label, measureRef.current.getBoundingClientRect().width)
  })

  const body = (
    <>
      <div className="xl-group-body">{children}</div>
      <div className="xl-group-label">{label}</div>
    </>
  )

  return (
    <div className="xl-group" style={style}>
      {scale && (
        <div ref={measureRef} className="xl-group xl-group-measure" aria-hidden="true">
          {body}
        </div>
      )}
      {collapsed ? (
        <button ref={anchorRef} className={'xl-group-collapsed' + (open ? ' active' : '')} onClick={() => setOpen((o) => !o)}>
          <span className="xl-big-icon">{icon}</span>
          <span className="xl-big-label">{label}<ChevronDown12Regular className="xl-drop-caret" /></span>
        </button>
      ) : body}
      {collapsed && open && (
        <GroupPopup anchor={anchorRef.current} onClose={() => setOpen(false)}>
          <div className="xl-group">{body}</div>
        </GroupPopup>
      )}
    </div>
  )
}

export function GroupSep() {
  return <div className="xl-group-sep" />
}

/** Editable combo box (font name / size / number format). */
export function Combo({
  value, options, onChange, onPick, width = 110, tip, editable = true, keytip
}: {
  value: string
  options: string[]
  onChange?: (v: string) => void
  onPick?: (v: string) => void
  width?: number
  tip?: Tip
  editable?: boolean
  keytip?: string
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
      <div className="xl-combo" ref={ref} style={{ width, position: 'relative' }}>
        <KeyTipBadge code={keytip} />
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

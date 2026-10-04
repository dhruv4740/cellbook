// Floats pictures and shapes over the grid: positioned from a cell anchor (scroll / zoom / frozen-pane aware), with
// Excel's selection frame, drag-to-move, resize handles, Delete, a right-click menu and double-click text editing.
import { useEffect, useRef, useState } from 'react'
import { useStore, SheetState } from '../model/store'
import { setUi, useUi } from '../chrome/ui'
import { PictureDef, ShapeDef } from '../model/types'
import type { ChartGeometry } from '../charts/ChartOverlay'
import { Menu, MenuItem } from '../components/Menu'
import { picturesOf, updatePicture, deletePicture } from '../data/pictures'
import { shapesOf, updateShape, deleteShape } from '../data/shapes'
import { floatsOf } from '../data/arrange'
import { arrangeOrder, commitMoveOfSelection, selectObject, selectObjects, selectedObjectIds } from '../data/arrangeActions'
import { FloatBox, Box } from './FloatBox'
import { Pt, addInk, eraseInks, pathD, pointInPolygon, screenPoints, strokeHit, truncate } from '../data/ink'

const FILLS: [string, string | undefined][] = [['Blue', '#4472c4'], ['Orange', '#ed7d31'], ['Gray', '#a5a5a5'], ['Gold', '#ffc000'], ['Green', '#70ad47'], ['Red', '#c00000'], ['White', '#ffffff'], ['No Fill', undefined]]
const LINES: [string, string | undefined][] = [['Dark Blue', '#2f528f'], ['Black', '#000000'], ['Red', '#c00000'], ['Green', '#548235'], ['Gray', '#7f7f7f'], ['No Outline', undefined]]

export function ObjectOverlay({ sheet, geo }: { sheet: SheetState; geo: ChartGeometry }) {
  useStore()
  const ui = useUi()
  const pics = picturesOf(sheet)
  const shapes = shapesOf(sheet)
  const selectedId = ui.selectedPicture // pictures and shapes share this selection (their ids never collide)
  const selIds = selectedObjectIds()
  const zOf = new Map(floatsOf(sheet.layout).map((f, i) => [f.id, i + 1]))
  const [band, setBand] = useState<{ x0: number; y0: number; x1: number; y1: number } | null>(null)
  const [menu, setMenu] = useState<{ x: number; y: number; id: string } | null>(null)
  const [editing, setEditing] = useState<string | null>(null)
  const [live, setLive] = useState<Pt[] | null>(null) // the stroke being drawn (host-relative px)
  const [erased, setErased] = useState<string[]>([])
  const { colsAxis, rowsAxis, frozen, scroll, z, rowHeaderW, headerH, viewport } = geo
  const frozenW = colsAxis.totalSize(frozen.cols) * z
  const frozenH = rowsAxis.totalSize(frozen.rows) * z

  useEffect(() => {
    if (!ui.selectObjectsMode && !ui.inkTool) return
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') setUi({ selectObjectsMode: false, inkTool: null }) }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [ui.selectObjectsMode, ui.inkTool])

  useEffect(() => {
    if (!selectedId || editing) return
    const onKey = (e: KeyboardEvent): void => {
      if (e.key !== 'Delete' && e.key !== 'Backspace') return
      const el = document.activeElement
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA')) return
      if (selectedId.startsWith('shp_')) deleteShape(sheet, selectedId); else deletePicture(sheet, selectedId)
      setUi({ selectedPicture: null })
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [selectedId, sheet, editing])

  if (!floatsOf(sheet.layout).length && !ui.inkTool) return null
  const sx = (c: number): number => (c < frozen.cols ? colsAxis.offset(c) * z : frozenW + colsAxis.offset(c) * z - colsAxis.offset(frozen.cols) * z - scroll.x)
  const sy = (r: number): number => (r < frozen.rows ? rowsAxis.offset(r) * z : frozenH + rowsAxis.offset(r) * z - rowsAxis.offset(frozen.rows) * z - scroll.y)
  const lx = (v: number): number => (v < frozenW ? v / z : colsAxis.offset(frozen.cols) + (v - frozenW + scroll.x) / z)
  const ly = (v: number): number => (v < frozenH ? v / z : rowsAxis.offset(frozen.rows) + (v - frozenH + scroll.y) / z)
  const boxOf = (p: { anchor: PictureDef['anchor']; width: number; height: number }): Box => ({ left: sx(p.anchor.c) + p.anchor.dx * z, top: sy(p.anchor.r) + p.anchor.dy * z, width: p.width * z, height: p.height * z })
  const place = (b: Box): { anchor: PictureDef['anchor']; width: number; height: number } => {
    const x = lx(b.left), y = ly(b.top)
    const c = colsAxis.indexAt(x), r = rowsAxis.indexAt(y)
    return { anchor: { r, c, dx: x - colsAxis.offset(c), dy: y - rowsAxis.offset(r) }, width: Math.max(8, b.width / z), height: Math.max(1, b.height / z) }
  }
  const select = (id: string, additive = false): void => selectObject(sheet, id, additive)
  const commit = (id: string, b: Box, update: () => void): void => {
    const pl = place(b)
    if (!commitMoveOfSelection(sheet, id, pl.anchor, pl.width, pl.height)) update()
  }
  const startBand = (e: React.MouseEvent<HTMLDivElement>): void => {
    const host = e.currentTarget.getBoundingClientRect()
    const x0 = e.clientX - host.left, y0 = e.clientY - host.top
    setBand({ x0, y0, x1: x0, y1: y0 })
    let last = { x1: x0, y1: y0 }
    const move = (ev: MouseEvent): void => { last = { x1: ev.clientX - host.left, y1: ev.clientY - host.top }; setBand({ x0, y0, ...last }) }
    const up = (): void => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      setBand(null)
      const l = Math.min(x0, last.x1), r = Math.max(x0, last.x1), t = Math.min(y0, last.y1), b = Math.max(y0, last.y1)
      const hit = floatsOf(sheet.layout).filter((f) => {
        if (f.obj.hidden) return false
        const bx = boxOf(f.obj)
        return bx.left < r && bx.left + bx.width > l && bx.top < b && bx.top + bx.height > t
      }).map((f) => f.id)
      selectObjects(sheet, hit)
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }

  /** Draw tab: pen strokes become ink objects, the eraser removes whole strokes, the lasso selects the objects it encloses. */
  const startInk = (e: React.MouseEvent<HTMLDivElement>): void => {
    const tool = ui.inkTool
    if (!tool || tool.kind === 'select') return
    e.preventDefault(); e.stopPropagation()
    const host = e.currentTarget.getBoundingClientRect()
    const at = (ev: { clientX: number; clientY: number }): Pt => [ev.clientX - host.left, ev.clientY - host.top]
    const pts: Pt[] = [at(e)]
    const gone = new Set<string>()
    const sweep = (p: Pt): void => {
      for (const s of shapes) if (s.kind === 'ink' && !s.hidden && !gone.has(s.id) && strokeHit(screenPoints(s, boxOf(s)), p, 6 + ((s.lineWidth ?? 1) * z) / 2)) gone.add(s.id)
      setErased([...gone])
    }
    if (tool.kind === 'eraser') sweep(pts[0])
    setLive([...pts])
    const move = (ev: MouseEvent): void => {
      const p = at(ev)
      pts.push(p)
      if (tool.kind === 'eraser') sweep(p)
      else setLive([...pts])
    }
    const up = (): void => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      setLive(null)
      setErased([])
      if (tool.kind === 'pen') addInk(sheet, pts.map((p) => [lx(p[0]), ly(p[1])] as Pt), tool.pen)
      else if (tool.kind === 'eraser') eraseInks(sheet, [...gone])
      else if (tool.kind === 'lasso') {
        selectObjects(sheet, floatsOf(sheet.layout).filter((f) => { if (f.obj.hidden) return false; const b = boxOf(f.obj); return pointInPolygon([b.left + b.width / 2, b.top + b.height / 2], pts) }).map((f) => f.id))
      }
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }
  /** Review > Hide Ink and Ink Replay decide which ink strokes show, and how much of each. */
  const inkFrac = (id: string): number => {
    const rp = ui.inkReplay
    if (!rp) return 1
    const i = rp.order.indexOf(id)
    return i < 0 || i < rp.shown ? 1 : i === rp.shown ? rp.frac : -1
  }

  const menuItems = (id: string): MenuItem[] => {
    const isShape = id.startsWith('shp_')
    const sh = shapes.find((s) => s.id === id)
    const items: MenuItem[] = [
      { label: 'Bring to &Front', onClick: () => arrangeOrder('front') },
      { label: 'S&end to Back', onClick: () => arrangeOrder('back') }
    ]
    if (sh && sh.kind !== 'ink') {
      const isLine = sh.kind === 'line' || sh.kind === 'arrow'
      items.push('separator', { label: 'Edit &Text', disabled: isLine, onClick: () => setEditing(id) })
      if (!isLine) {
        items.push('separator', { label: 'Shape Fill:', disabled: true })
        for (const [name, col] of FILLS) items.push({ label: name, checked: sh.fill === col, onClick: () => updateShape(sheet, id, { fill: col, ...(col && sh.kind === 'textbox' ? { color: sh.color ?? '#000000' } : {}) }, 'Shape Fill') })
      }
      items.push('separator', { label: 'Shape Outline:', disabled: true })
      for (const [name, col] of LINES) items.push({ label: name, checked: sh.line === col, onClick: () => updateShape(sheet, id, { line: col, lineWidth: sh.lineWidth ?? 1 }, 'Shape Outline') })
    }
    items.push('separator', { label: '&Delete', onClick: () => { if (isShape) deleteShape(sheet, id); else deletePicture(sheet, id); setUi({ selectedPicture: null }) } })
    return items
  }

  return (
    <div style={{ position: 'absolute', left: rowHeaderW, top: headerH, width: Math.max(0, viewport.w - rowHeaderW), height: Math.max(0, viewport.h - headerH), overflow: 'hidden', pointerEvents: 'none' }}>
      {ui.selectObjectsMode && <div onMouseDown={startBand} style={{ position: 'absolute', inset: 0, pointerEvents: 'auto', cursor: 'crosshair' }} />}
      {ui.inkTool && ui.inkTool.kind !== 'select' && (
        <div onMouseDown={startInk} style={{ position: 'absolute', inset: 0, pointerEvents: 'auto', cursor: 'crosshair', zIndex: 200000 }}>
          {live && (
            <svg width="100%" height="100%" style={{ position: 'absolute', inset: 0, pointerEvents: 'none' }}>
              {ui.inkTool.kind === 'pen' && <path d={pathD(live)} fill="none" stroke={ui.inkTool.pen.color} strokeOpacity={ui.inkTool.pen.alpha ?? 1} strokeWidth={ui.inkTool.pen.width * z} strokeLinecap="round" strokeLinejoin="round" />}
              {ui.inkTool.kind === 'lasso' && <path d={pathD(live) + ' Z'} fill="rgba(33,115,70,0.08)" stroke="#217346" strokeDasharray="4 3" />}
            </svg>
          )}
        </div>
      )}
      {band && <div style={{ position: 'absolute', left: Math.min(band.x0, band.x1), top: Math.min(band.y0, band.y1), width: Math.abs(band.x1 - band.x0), height: Math.abs(band.y1 - band.y0), border: '1px dashed #217346', background: 'rgba(33,115,70,0.08)', zIndex: 100000 }} />}
      {pics.filter((p) => !p.hidden).map((p) => (
        <FloatBox key={p.id} box={boxOf(p)} lockAspect zIndex={zOf.get(p.id)} rotation={p.rotation} selected={selIds.includes(p.id)} onSelect={(a) => select(p.id, a)}
          onContext={(x, y) => { select(p.id); setMenu({ x, y, id: p.id }) }} onCommit={(b) => commit(p.id, b, () => updatePicture(sheet, p.id, place(b), 'Move or Resize Picture'))}>
          <img src={p.src} alt={p.alt ?? p.name} draggable={false} style={{ width: '100%', height: '100%', display: 'block', userSelect: 'none', transform: flipCss(p) }} />
        </FloatBox>
      ))}
      {shapes.filter((s) => !s.hidden && !(s.kind === 'ink' && (ui.hideInk || erased.includes(s.id) || inkFrac(s.id) < 0))).map((s) => (
        <FloatBox key={s.id} box={boxOf(s)} minSize={s.kind === 'line' || s.kind === 'arrow' ? 1 : 16} zIndex={zOf.get(s.id)} rotation={s.rotation} selected={selIds.includes(s.id)} onSelect={(a) => select(s.id, a)}
          onDoubleClick={() => { if (s.kind !== 'line' && s.kind !== 'arrow' && s.kind !== 'ink') setEditing(s.id) }}
          onContext={(x, y) => { select(s.id); setMenu({ x, y, id: s.id }) }} onCommit={(b) => commit(s.id, b, () => updateShape(sheet, s.id, place(b), 'Move or Resize Shape'))}>
          <ShapeView shape={s} z={z} box={boxOf(s)} frac={s.kind === 'ink' ? inkFrac(s.id) : 1} editing={editing === s.id} onText={(t) => { updateShape(sheet, s.id, { text: t || undefined }, 'Edit Text'); setEditing(null) }} />
        </FloatBox>
      ))}
      {menu && (
        <div style={{ pointerEvents: 'auto' }}>
          <Menu x={menu.x} y={menu.y} onClose={() => setMenu(null)} items={menuItems(menu.id)} />
        </div>
      )}
    </div>
  )
}

/** CSS flip for the picture or a shape's outline (text stays upright). */
const flipCss = (o: { flipH?: boolean; flipV?: boolean }): string | undefined => (o.flipH || o.flipV ? `scale(${o.flipH ? -1 : 1}, ${o.flipV ? -1 : 1})` : undefined)

function ShapeView({ shape: s, z, box, frac, editing, onText }: { shape: ShapeDef; z: number; box: Box; frac: number; editing: boolean; onText: (t: string) => void }) {
  const w = box.width, h = box.height
  const fill = s.fill ?? 'none'
  const stroke = s.line ?? 'none'
  const sw = (s.lineWidth ?? 1) * z
  const geom = ((): React.ReactElement => {
    switch (s.kind) {
      case 'ellipse': return <ellipse cx={w / 2} cy={h / 2} rx={Math.max(0, w / 2 - sw / 2)} ry={Math.max(0, h / 2 - sw / 2)} fill={fill} stroke={stroke} strokeWidth={sw} />
      case 'triangle': return <polygon points={`${w / 2},${sw / 2} ${w - sw / 2},${h - sw / 2} ${sw / 2},${h - sw / 2}`} fill={fill} stroke={stroke} strokeWidth={sw} />
      case 'diamond': return <polygon points={`${w / 2},${sw / 2} ${w - sw / 2},${h / 2} ${w / 2},${h - sw / 2} ${sw / 2},${h / 2}`} fill={fill} stroke={stroke} strokeWidth={sw} />
      case 'rightArrow': {
        const nk = h * 0.25, hx = Math.max(w - h * 0.5, w * 0.5)
        return <polygon points={`0,${nk} ${hx},${nk} ${hx},0 ${w},${h / 2} ${hx},${h} ${hx},${h - nk} 0,${h - nk}`} fill={fill} stroke={stroke} strokeWidth={sw} />
      }
      case 'line': case 'arrow': {
        const y1 = s.flipV ? h : 0, y2 = s.flipV ? 0 : h
        return (
          <>
            <defs><marker id={'ah' + s.id} markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto-start-reverse"><path d="M0,0 L8,4 L0,8 z" fill={stroke} /></marker></defs>
            <line x1={0} y1={y1} x2={w} y2={y2} stroke={stroke} strokeWidth={sw} markerEnd={s.kind === 'arrow' ? `url(#ah${s.id})` : undefined} />
          </>
        )
      }
      case 'ink': {
        const pts = truncate(screenPoints(s, { left: 0, top: 0, width: w, height: h }), frac)
        return pts.length === 1
          ? <circle cx={pts[0][0]} cy={pts[0][1]} r={sw / 2} fill={stroke} fillOpacity={s.alpha ?? 1} />
          : <path d={pathD(pts)} fill="none" stroke={stroke} strokeOpacity={s.alpha ?? 1} strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" />
      }
      case 'roundRect': return <rect x={sw / 2} y={sw / 2} width={Math.max(0, w - sw)} height={Math.max(0, h - sw)} rx={Math.min(w, h) * 0.1667} fill={fill} stroke={stroke} strokeWidth={sw} />
      default: return <rect x={sw / 2} y={sw / 2} width={Math.max(0, w - sw)} height={Math.max(0, h - sw)} fill={fill} stroke={stroke} strokeWidth={sw} />
    }
  })()
  const align = s.align === 'r' ? 'flex-end' : s.align === 'ctr' ? 'center' : 'flex-start'
  const textStyle: React.CSSProperties = {
    color: s.color ?? '#000000', fontSize: (s.fontSize ?? 11) * z * (96 / 72), fontWeight: s.bold ? 700 : 400, fontStyle: s.italic ? 'italic' : 'normal',
    textAlign: s.align === 'r' ? 'right' : s.align === 'ctr' ? 'center' : 'left', fontFamily: s.fontFamily ? `"${s.fontFamily}", Calibri, Carlito, sans-serif` : 'Calibri, Carlito, sans-serif', whiteSpace: 'pre-wrap', overflow: 'hidden', width: '100%',
    ...(s.textOutline ? { WebkitTextStroke: `${Math.max(0.5, z)}px ${s.textOutline}`, paintOrder: 'stroke fill' } : {}),
    ...(s.textShadow ? { textShadow: `${2 * z}px ${2 * z}px ${3 * z}px rgba(0,0,0,0.45)` } : {})
  }
  const isLine = s.kind === 'line' || s.kind === 'arrow' || s.kind === 'ink'
  return (
    <>
      <svg width={w} height={h} style={{ position: 'absolute', left: 0, top: 0, overflow: 'visible', pointerEvents: isLine ? 'none' : 'auto', transform: isLine && s.kind !== 'ink' ? undefined : flipCss(s) }}>{geom}</svg>
      {isLine && <div style={{ position: 'absolute', inset: -4, pointerEvents: 'auto' }} />}
      {!isLine && (
        <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: s.kind === 'textbox' ? 'flex-start' : 'center', justifyContent: align, padding: 4 * z, pointerEvents: editing ? 'auto' : 'none' }}>
          {editing ? <ShapeEditor text={s.text ?? ''} style={textStyle} onDone={onText} /> : <div style={textStyle}>{s.text}</div>}
        </div>
      )}
    </>
  )
}

function ShapeEditor({ text, style, onDone }: { text: string; style: React.CSSProperties; onDone: (t: string) => void }) {
  const ref = useRef<HTMLTextAreaElement>(null)
  useEffect(() => { ref.current?.focus(); ref.current?.select() }, [])
  return (
    <textarea ref={ref} defaultValue={text} spellCheck={false}
      style={{ ...style, height: '100%', border: 'none', outline: 'none', background: 'rgba(255,255,255,0.15)', resize: 'none' }}
      onMouseDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Escape') onDone(text) }}
      onBlur={(e) => onDone(e.currentTarget.value)} />
  )
}

// Floats chart objects over the grid: positions them from cell anchors (scrolling/zoom/frozen-pane
// aware), and gives the selected chart Excel's selection frame, drag-move, resize handles, the
// right-click menu and the three floating buttons (Chart Elements / Styles / Filters).
import { useEffect, useMemo, useRef, useState } from 'react'
import { useStore, SheetState } from '../model/store'
import { setUi, useUi } from '../chrome/ui'
import { ChartSpec, Range, rangeA1 } from '../model/types'
import { Axis } from '../grid/axis'
import { extractChartData } from './dataExtraction'
import { ChartRenderer } from './ChartRenderer'
import { chartsOf, readerFor, deleteChart, updateChart, moveOrResizeChart, bringChartToFront, sendChartToBack } from './chartsStore'
import { Menu, MenuItem } from '../components/Menu'
import {
  Add20Regular, PaintBrush20Regular, Filter20Regular
} from '@fluentui/react-icons'

export interface ChartGeometry {
  colsAxis: Axis
  rowsAxis: Axis
  frozen: { rows: number; cols: number }
  scroll: { x: number; y: number }
  z: number
  rowHeaderW: number
  headerH: number
  viewport: { w: number; h: number }
}

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
type Handle = (typeof HANDLES)[number]

export default function ChartOverlay({ sheet, geo }: { sheet: SheetState; geo: ChartGeometry }) {
  const store = useStore()
  const ui = useUi()
  const charts = chartsOf(sheet)
  const selectedId = ui.selectedChart
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; id: string } | null>(null)
  const [panel, setPanel] = useState<null | 'elements' | 'styles' | 'filters'>(null)

  const { colsAxis, rowsAxis, frozen, scroll, z, rowHeaderW, headerH, viewport } = geo
  const frozenW = colsAxis.totalSize(frozen.cols) * z
  const frozenH = rowsAxis.totalSize(frozen.rows) * z

  // Delete removes the selected chart, as long as nothing else (a cell edit, a text field) has focus.
  useEffect(() => {
    if (!selectedId) return
    const onKey = (e: KeyboardEvent): void => {
      if (e.key !== 'Delete' && e.key !== 'Backspace') return
      const el = document.activeElement
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA')) return
      deleteChart(store, sheet, selectedId)
      setUi({ selectedChart: null })
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [selectedId, store, sheet])

  function screenXRel(c: number): number {
    return c < frozen.cols ? colsAxis.offset(c) * z : frozenW + colsAxis.offset(c) * z - colsAxis.offset(frozen.cols) * z - scroll.x
  }
  function screenYRel(r: number): number {
    return r < frozen.rows ? rowsAxis.offset(r) * z : frozenH + rowsAxis.offset(r) * z - rowsAxis.offset(frozen.rows) * z - scroll.y
  }
  function toLogicalX(screenRel: number): number {
    return screenRel < frozenW ? screenRel / z : colsAxis.offset(frozen.cols) + (screenRel - frozenW + scroll.x) / z
  }
  function toLogicalY(screenRel: number): number {
    return screenRel < frozenH ? screenRel / z : rowsAxis.offset(frozen.rows) + (screenRel - frozenH + scroll.y) / z
  }
  function boxOf(spec: ChartSpec): { left: number; top: number; width: number; height: number } {
    return {
      left: screenXRel(spec.anchor.c) + spec.anchor.dx * z,
      top: screenYRel(spec.anchor.r) + spec.anchor.dy * z,
      width: spec.width * z,
      height: spec.height * z
    }
  }
  function commitBox(id: string, box: { left: number; top: number; width: number; height: number }): void {
    const lx = toLogicalX(box.left), ly = toLogicalY(box.top)
    const c = colsAxis.indexAt(lx), r = rowsAxis.indexAt(ly)
    const anchor = { r, c, dx: lx - colsAxis.offset(c), dy: ly - rowsAxis.offset(r) }
    moveOrResizeChart(store, sheet, id, anchor, Math.max(40, box.width / z), Math.max(30, box.height / z))
  }

  return (
    <div
      style={{ position: 'absolute', left: rowHeaderW, top: headerH, width: Math.max(0, viewport.w - rowHeaderW), height: Math.max(0, viewport.h - headerH), overflow: 'hidden', pointerEvents: 'none' }}
      onMouseDownCapture={() => panel && setPanel(null)}
    >
      {charts.map((spec) => {
        const box = boxOf(spec)
        const selected = spec.id === selectedId
        return (
          <ChartBox
            key={spec.id}
            spec={spec}
            box={box}
            z={z}
            selected={selected}
            sheet={sheet}
            onSelect={() => setUi({ selectedChart: spec.id })}
            onContextMenu={(x, y) => { setUi({ selectedChart: spec.id }); setCtxMenu({ x, y, id: spec.id }) }}
            onCommitBox={(b) => commitBox(spec.id, b)}
            panel={selected ? panel : null}
            onPanel={setPanel}
          />
        )
      })}
      {ctxMenu && (
        <div style={{ pointerEvents: 'auto' }}>
          <Menu x={ctxMenu.x} y={ctxMenu.y} onClose={() => setCtxMenu(null)} items={contextMenuItems(store, sheet, ctxMenu.id)} />
        </div>
      )}
    </div>
  )
}

function contextMenuItems(store: ReturnType<typeof useStore>, sheet: SheetState, id: string): MenuItem[] {
  return [
    { label: 'Cu&t', shortcut: 'Ctrl+X' },
    { label: '&Copy', shortcut: 'Ctrl+C' },
    { label: 'Paste &Options:' },
    'separator',
    { label: '&Reset to Match Style' },
    { label: '&Font...' },
    { label: 'Change C&hart Type...' },
    { label: 'Save as &Template...' },
    'separator',
    { label: 'Select &Data...' },
    { label: '&Move Chart...' },
    { label: '&3-D Rotation...' },
    'separator',
    { label: '&Group' },
    { label: 'Bring to &Front', onClick: () => bringChartToFront(store, sheet, id) },
    { label: 'S&end to Back', onClick: () => sendChartToBack(store, sheet, id) },
    'separator',
    { label: 'Assign &Macro...' },
    { label: 'Edit Alt &Text...' },
    { label: 'For&mat Chart Area...' },
    'separator',
    { label: '&Delete', onClick: () => { deleteChart(store, sheet, id); setUi({ selectedChart: null }) } }
  ]
}

function ChartBox({ spec, box, z, selected, sheet, onSelect, onContextMenu, onCommitBox, panel, onPanel }: {
  spec: ChartSpec
  box: { left: number; top: number; width: number; height: number }
  z: number
  selected: boolean
  sheet: SheetState
  onSelect: () => void
  onContextMenu: (x: number, y: number) => void
  onCommitBox: (box: { left: number; top: number; width: number; height: number }) => void
  panel: null | 'elements' | 'styles' | 'filters'
  onPanel: (p: null | 'elements' | 'styles' | 'filters') => void
}) {
  const store = useStore()
  const [drag, setDrag] = useState<{ box: typeof box } | null>(null)
  const startRef = useRef<{ mx: number; my: number; box: typeof box; handle?: Handle } | null>(null)

  const liveBox = drag?.box ?? box
  const data = useMemo(() => {
    const filtered = extractChartData(readerFor(store, sheet), spec.dataRange, spec.seriesInColumns)
    const hide = new Set(spec.hiddenSeriesIdx ?? [])
    return { categories: filtered.categories, series: filtered.series.filter((_, i) => !hide.has(i)) }
  }, [store, sheet, spec])

  function beginDrag(e: React.MouseEvent, handle?: Handle): void {
    e.preventDefault()
    e.stopPropagation()
    onSelect()
    startRef.current = { mx: e.clientX, my: e.clientY, box, handle }
    const move = (ev: MouseEvent): void => {
      const s = startRef.current
      if (!s) return
      const dx = ev.clientX - s.mx, dy = ev.clientY - s.my
      let b = { ...s.box }
      if (!s.handle) {
        b.left = Math.max(0, s.box.left + dx)
        b.top = Math.max(0, s.box.top + dy)
      } else {
        if (s.handle.includes('e')) b.width = Math.max(40, s.box.width + dx)
        if (s.handle.includes('s')) b.height = Math.max(30, s.box.height + dy)
        if (s.handle.includes('w')) { b.width = Math.max(40, s.box.width - dx); b.left = s.box.left + (s.box.width - b.width) }
        if (s.handle.includes('n')) { b.height = Math.max(30, s.box.height - dy); b.top = s.box.top + (s.box.height - b.height) }
      }
      setDrag({ box: b })
    }
    const up = (): void => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      startRef.current = null
      setDrag((cur) => {
        if (cur) onCommitBox(cur.box)
        return null
      })
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }

  return (
    <div
      style={{
        position: 'absolute', left: liveBox.left, top: liveBox.top, width: liveBox.width, height: liveBox.height,
        pointerEvents: 'auto', boxShadow: selected ? 'none' : '0 0 0 1px rgba(0,0,0,0.06)',
        outline: selected ? '1px solid #217346' : 'none', background: '#fff', cursor: drag ? 'move' : 'default'
      }}
      onMouseDown={(e) => beginDrag(e)}
      onClick={onSelect}
      onContextMenu={(e) => { e.preventDefault(); onContextMenu(e.clientX, e.clientY) }}
    >
      <ChartRenderer spec={spec} data={data} width={Math.max(1, liveBox.width / z)} height={Math.max(1, liveBox.height / z)} />
      {selected && (
        <>
          {HANDLES.map((h) => (
            <div
              key={h}
              onMouseDown={(e) => beginDrag(e, h)}
              style={{ position: 'absolute', width: 7, height: 7, background: '#fff', border: '1px solid #217346', borderRadius: '50%', ...handlePos(h) }}
            />
          ))}
          <FloatingButtons spec={spec} sheet={sheet} panel={panel} onPanel={onPanel} />
        </>
      )}
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

function FloatingButtons({ spec, sheet, panel, onPanel }: { spec: ChartSpec; sheet: SheetState; panel: null | 'elements' | 'styles' | 'filters'; onPanel: (p: null | 'elements' | 'styles' | 'filters') => void }) {
  const store = useStore()
  const btn = (icon: React.ReactNode, key: 'elements' | 'styles' | 'filters', title: string): JSX.Element => (
    <button
      title={title}
      onMouseDown={(e) => e.stopPropagation()}
      onClick={(e) => { e.stopPropagation(); onPanel(panel === key ? null : key) }}
      style={{
        width: 24, height: 24, display: 'flex', alignItems: 'center', justifyContent: 'center',
        background: panel === key ? '#E8F1FB' : '#fff', border: '1px solid #C6C6C6', borderRadius: 3, cursor: 'pointer', color: '#217346'
      }}
    >
      {icon}
    </button>
  )
  return (
    <div style={{ position: 'absolute', left: '100%', top: 0, marginLeft: 6, display: 'flex', flexDirection: 'column', gap: 6 }} onMouseDown={(e) => e.stopPropagation()}>
      {btn(<Add20Regular fontSize={16} />, 'elements', 'Chart Elements')}
      {btn(<PaintBrush20Regular fontSize={16} />, 'styles', 'Chart Styles')}
      {btn(<Filter20Regular fontSize={16} />, 'filters', 'Chart Filters')}
      {panel === 'elements' && <ElementsPanel spec={spec} sheet={sheet} store={store} />}
      {panel === 'styles' && <StylesPanel spec={spec} sheet={sheet} store={store} />}
      {panel === 'filters' && <FiltersPanel spec={spec} sheet={sheet} store={store} />}
    </div>
  )
}

function panelBox(): React.CSSProperties {
  return { position: 'absolute', left: 30, top: 0, width: 190, background: '#fff', border: '1px solid #C6C6C6', boxShadow: '0 2px 8px rgba(0,0,0,0.18)', borderRadius: 3, padding: 6, fontSize: 12, zIndex: 10 }
}

function ElementsPanel({ spec, sheet, store }: { spec: ChartSpec; sheet: SheetState; store: ReturnType<typeof useStore> }) {
  const rows: { key: keyof ChartSpec['elements']; label: string }[] = [
    { key: 'axes', label: 'Axes' },
    { key: 'axisTitles', label: 'Axis Titles' },
    { key: 'title', label: 'Chart Title' },
    { key: 'dataLabels', label: 'Data Labels' },
    { key: 'gridlines', label: 'Gridlines' },
    { key: 'legend', label: 'Legend' }
  ]
  return (
    <div style={panelBox()}>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>Chart Elements</div>
      {rows.map((row) => (
        <label key={row.key} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '3px 2px', cursor: 'pointer' }}>
          <input
            type="checkbox"
            checked={spec.elements[row.key]}
            onChange={() => updateChart(store, sheet, spec.id, { elements: { ...spec.elements, [row.key]: !spec.elements[row.key] } }, 'Chart Elements')}
          />
          {row.label}
        </label>
      ))}
    </div>
  )
}

function StylesPanel({ spec, sheet, store }: { spec: ChartSpec; sheet: SheetState; store: ReturnType<typeof useStore> }) {
  return (
    <div style={panelBox()}>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>Chart Styles</div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 6 }}>
        {Array.from({ length: 8 }, (_, i) => i).map((i) => (
          <button
            key={i}
            onClick={() => updateChart(store, sheet, spec.id, { styleId: i }, 'Chart Style')}
            style={{ width: 34, height: 24, border: (spec.styleId ?? 0) === i ? '2px solid #217346' : '1px solid #C6C6C6', borderRadius: 2, cursor: 'pointer', background: '#fafafa' }}
          />
        ))}
      </div>
      <div style={{ fontWeight: 600, margin: '8px 0 4px' }}>Color</div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
        {Array.from({ length: 7 }, (_, i) => i).map((i) => (
          <button
            key={i}
            title={i === 0 ? 'Colorful' : 'Monochromatic ' + i}
            onClick={() => updateChart(store, sheet, spec.id, { colorSetId: i }, 'Change Colors')}
            style={{ width: 20, height: 20, border: (spec.colorSetId ?? 0) === i ? '2px solid #217346' : '1px solid #C6C6C6', borderRadius: '50%', cursor: 'pointer', background: i === 0 ? 'conic-gradient(#156082,#E97132,#196B24,#0F9ED5,#A02B93,#4EA72E,#156082)' : '#5B9BD5' }}
          />
        ))}
      </div>
    </div>
  )
}

function FiltersPanel({ spec, sheet, store }: { spec: ChartSpec; sheet: SheetState; store: ReturnType<typeof useStore> }) {
  const data = useMemo(() => extractChartData(readerFor(store, sheet), spec.dataRange, spec.seriesInColumns), [store, sheet, spec])
  const hidden = new Set(spec.hiddenSeriesIdx ?? [])
  const toggle = (i: number): void => {
    const next = new Set(hidden)
    if (next.has(i)) next.delete(i)
    else next.add(i)
    updateChart(store, sheet, spec.id, { hiddenSeriesIdx: [...next] }, 'Chart Filters')
  }
  return (
    <div style={panelBox()}>
      <div style={{ fontWeight: 600, marginBottom: 4 }}>VALUES</div>
      {data.series.map((s, i) => (
        <label key={i} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '2px 2px', cursor: 'pointer' }}>
          <input type="checkbox" checked={!hidden.has(i)} onChange={() => toggle(i)} />
          {s.name}
        </label>
      ))}
      <div style={{ fontWeight: 600, margin: '6px 0 2px' }}>CATEGORIES</div>
      {data.categories.slice(0, 8).map((c, i) => (
        <label key={i} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '2px 2px', color: '#595959' }}>
          <input type="checkbox" checked disabled />
          {c}
        </label>
      ))}
    </div>
  )
}

export function rangeLabel(range: Range): string {
  return rangeA1(range)
}

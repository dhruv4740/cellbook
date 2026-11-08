// Floats slicers and timelines over the grid (same cell-anchored placement as ObjectOverlay).
import { useState } from 'react'
import { useStore, SheetState } from '../model/store'
import { SlicerDef } from '../model/types'
import type { ChartGeometry } from '../charts/ChartOverlay'
import { Menu } from '../components/Menu'
import { BLANK_TOKEN } from '../engine/filter'
import { FloatBox, Box } from './FloatBox'
import { slicersOf, slicerItems, timelineSpan, timelinePeriods, pickSlicerItem, pickTimelinePeriod, clearSlicer, deleteSlicer, updateSlicer, TimelinePeriod } from '../data/slicers'

const PERIODS: [TimelinePeriod, string][] = [['years', 'Years'], ['quarters', 'Quarters'], ['months', 'Months'], ['days', 'Days']]
const BLUE = '#4472c4'

export function SlicerOverlay({ sheet, geo }: { sheet: SheetState; geo: ChartGeometry }) {
  useStore()
  const [sel, setSel] = useState<string | null>(null)
  const [menu, setMenu] = useState<{ x: number; y: number; id: string } | null>(null)
  const list = slicersOf(sheet)
  if (!list.length) return null
  const { colsAxis, rowsAxis, frozen, scroll, z } = geo
  const frozenW = colsAxis.totalSize(frozen.cols) * z
  const frozenH = rowsAxis.totalSize(frozen.rows) * z
  const sx = (c: number): number => (c < frozen.cols ? colsAxis.offset(c) * z : frozenW + colsAxis.offset(c) * z - colsAxis.offset(frozen.cols) * z - scroll.x)
  const sy = (r: number): number => (r < frozen.rows ? rowsAxis.offset(r) * z : frozenH + rowsAxis.offset(r) * z - rowsAxis.offset(frozen.rows) * z - scroll.y)
  const lx = (v: number): number => (v < frozenW ? v / z : colsAxis.offset(frozen.cols) + (v - frozenW + scroll.x) / z)
  const ly = (v: number): number => (v < frozenH ? v / z : rowsAxis.offset(frozen.rows) + (v - frozenH + scroll.y) / z)
  const boxOf = (s: SlicerDef): Box => ({ left: sx(s.anchor.c) + s.anchor.dx * z, top: sy(s.anchor.r) + s.anchor.dy * z, width: s.width * z, height: s.height * z })
  const commit = (s: SlicerDef, b: Box): void => {
    const x = lx(b.left), y = ly(b.top), c = colsAxis.indexAt(x), r = rowsAxis.indexAt(y)
    updateSlicer(sheet, s.id, { anchor: { r, c, dx: x - colsAxis.offset(c), dy: y - rowsAxis.offset(r) }, width: Math.max(80, b.width / z), height: Math.max(60, b.height / z) }, 'Move Slicer')
  }

  return (
    <>
      {list.map((s) => (
        <FloatBox key={s.id} box={boxOf(s)} selected={sel === s.id} minSize={60} zIndex={30}
          onSelect={() => setSel(s.id)} onContext={(x, y) => { setSel(s.id); setMenu({ x, y, id: s.id }) }} onCommit={(b) => commit(s, b)}>
          <div className="xl-slicer" data-slicer={s.name} style={{ position: 'absolute', inset: 0, background: '#fff', border: '1px solid #b3b0ad', display: 'flex', flexDirection: 'column', font: `${12 * z}px var(--ui-font)`, overflow: 'hidden' }}>
            <div style={{ display: 'flex', alignItems: 'center', padding: `${3 * z}px ${6 * z}px`, background: '#f3f2f1', borderBottom: '1px solid #d1d1d1' }}>
              <b style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.kind === 'timeline' ? s.header : s.name}</b>
              {s.kind === 'timeline' && (
                <select value={s.period ?? 'months'} onMouseDown={(e) => e.stopPropagation()} onChange={(e) => { updateSlicer(sheet, s.id, { period: e.target.value as TimelinePeriod, range: null }); clearSlicer(sheet, { ...s, period: e.target.value as TimelinePeriod }) }} style={{ fontSize: 11 * z, marginRight: 4 }}>
                  {PERIODS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                </select>
              )}
              <button title="Clear Filter" aria-label="Clear Filter" onMouseDown={(e) => e.stopPropagation()} onClick={() => clearSlicer(sheet, s)} style={{ border: 0, background: 'transparent', cursor: 'pointer', fontSize: 12 * z }}>✕</button>
            </div>
            {s.kind === 'slicer' ? <ItemList sheet={sheet} s={s} z={z} /> : <Timeline sheet={sheet} s={s} z={z} />}
          </div>
        </FloatBox>
      ))}
      {menu && <Menu x={menu.x} y={menu.y} onClose={() => setMenu(null)} items={[
        { label: 'Clear Filter', onClick: () => { const s = list.find((x) => x.id === menu.id); if (s) clearSlicer(sheet, s) } },
        { label: s_label(list, menu.id), onClick: () => deleteSlicer(sheet, menu.id) }
      ]} />}
    </>
  )
}

const s_label = (list: SlicerDef[], id: string): string => `Remove ${list.find((x) => x.id === id)?.kind === 'timeline' ? 'Timeline' : 'Slicer'}`

function ItemList({ sheet, s, z }: { sheet: SheetState; s: SlicerDef; z: number }) {
  const items = slicerItems(sheet, s)
  return (
    <div style={{ flex: 1, overflowY: 'auto', padding: 4 * z, display: 'flex', flexDirection: 'column', gap: 3 * z }} onMouseDown={(e) => e.stopPropagation()}>
      {items.map((it) => {
        const on = !s.selected || s.selected.includes(it)
        return (
          <div key={it} role="button" data-slicer-item={it} onClick={(e) => pickSlicerItem(sheet, s, it, e.ctrlKey || e.metaKey)}
            style={{ padding: `${3 * z}px ${8 * z}px`, cursor: 'pointer', userSelect: 'none', border: `1px solid ${on ? BLUE : '#d1d1d1'}`, background: on ? '#dce6f4' : '#f7f7f7', color: on ? '#000' : '#8a8886' }}>
            {it === BLANK_TOKEN ? '(blank)' : it}
          </div>
        )
      })}
    </div>
  )
}

function Timeline({ sheet, s, z }: { sheet: SheetState; s: SlicerDef; z: number }) {
  const span = timelineSpan(sheet, s)
  const periods = span ? timelinePeriods(span, s.period ?? 'months') : []
  return (
    <div style={{ flex: 1, overflowX: 'auto', display: 'flex', alignItems: 'center', gap: 2 * z, padding: 6 * z }} onMouseDown={(e) => e.stopPropagation()}>
      {periods.map((p) => {
        const on = !!s.range && p.lo >= s.range[0] && p.hi <= s.range[1]
        return (
          <div key={p.lo} role="button" data-timeline-period={p.label} onClick={(e) => pickTimelinePeriod(sheet, s, p, e.shiftKey)}
            style={{ flex: '0 0 auto', minWidth: 44 * z, textAlign: 'center', padding: `${14 * z}px ${6 * z}px`, cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap', border: `1px solid ${on ? BLUE : '#d1d1d1'}`, background: on ? BLUE : '#f7f7f7', color: on ? '#fff' : '#242424' }}>
            {p.label}
          </div>
        )
      })}
      {!periods.length && <span style={{ color: '#8a8886' }}>No dates</span>}
    </div>
  )
}

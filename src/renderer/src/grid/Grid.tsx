// The worksheet grid: canvas-rendered, virtualized, with Excel's exact selection/keyboard/editing
// behavior. Cell contents & formulas live in HyperFormula (via the store); this file only draws and
// interacts with them.
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { store, useStore, SheetState } from '../model/store'
import { getUi, setUi, useUi, commitEdit, cancelEdit, commitEditToSelection } from '../chrome/ui'
import {
  CellStyle, Range, CellRef, Selection, normRange, inRange, colName, a1, rangeA1, parseA1,
  DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, MAX_ROWS, MAX_COLS
} from '../model/types'
import { buildAxis, Axis } from './axis'
import { outlineRuns, maxLevel } from '../model/outline'
import { drawGutter, hitGutter, gutterSizes, GutterGeo } from './outlineGutter'
import { showOutlineLevel, toggleRun, outlineCommand } from './outlineCommands'
import { Menu, MenuItem, useMenu } from '../components/Menu'
import { Dialog } from '../components/Dialog'
import * as cmd from './commands'
import { shiftFormula } from '../engine/shiftFormula'
import ChartOverlay from '../charts/ChartOverlay'
import { insertChartFromSelection } from '../charts/insertChart'
import { COLUMN_BAR_GALLERY } from '../charts/gallery'
import { AutoFilterMenu } from './AutoFilterMenu'
import { NoteBox, NoteTooltip, CommentBox, closeNoteEditor, closeCommentPanel } from './Annotations'
import { evaluateCf } from '../format/conditional'
import { linkAt, followLink } from './links'
import { renderPivot } from '../data/pivot'
import { getValidationForCell, getListItems } from '../data/dataValidation'
import {
  Cut20Regular, Copy20Regular, ClipboardPaste20Regular, ClipboardTaskRegular, ClipboardCodeRegular,
  ClipboardLink20Regular, ArrowSwap20Regular, PaintBrush20Regular, Delete20Regular, Eraser20Regular,
  TableAdd20Regular, TableDeleteRow20Regular, ArrowSort20Regular, Filter20Regular, Comment20Regular,
  Note20Regular, Link20Regular, TableCellEdit20Regular, Translate20Regular, SquareHintSparkles20Regular,
  ColumnTriple20Regular, TextBold20Regular, TextItalic20Regular, TextUnderline20Regular,
  AlignLeft20Regular, AlignCenterHorizontal20Regular, AlignRight20Regular, BorderAll20Regular,
  PaintBucket20Regular, TextFont20Regular, FontIncrease20Regular, FontDecrease20Regular,
  TextPercent20Regular, TextIndentDecrease20Regular, TextIndentIncrease20Regular
} from '@fluentui/react-icons'

const HEADER_H = 20
const CORNER_W_MIN = 46
const REF_COLORS = ['#2563eb', '#dc2626', '#7c3aed', '#059669', '#d97706', '#0891b2', '#be185d', '#4d7c0f']

function fontString(style: CellStyle | undefined, z: number): string {
  const size = Math.round((style?.fontSize ?? 11) * z)
  const fam = style?.fontFamily ?? 'Calibri'
  let s = ''
  if (style?.italic) s += 'italic '
  if (style?.bold) s += 'bold '
  return `${s}${size}px "${fam}", "Carlito", sans-serif`
}

interface Pt { x: number; y: number }

function MiniBtn({ icon, title, active, onClick }: { icon: React.ReactNode; title: string; active?: boolean; onClick: () => void }) {
  return (
    <button
      type="button"
      className={'xl-mini-btn' + (active ? ' active' : '')}
      title={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {icon}
    </button>
  )
}

function drawCfIcon(ctx: CanvasRenderingContext2D, set: string, index: number, x: number, y: number, z: number): void {
  ctx.save()
  const r = 5 * z
  if (set.includes('TrafficLights') || set.includes('Shapes') || set.includes('Symbols')) {
    const colors = ['#e81123', '#ffb900', '#107c41']
    ctx.fillStyle = colors[index] || '#107c41'
    ctx.beginPath()
    ctx.arc(x + r, y, r, 0, 2 * Math.PI)
    ctx.fill()
    ctx.strokeStyle = '#333'
    ctx.lineWidth = 0.5
    ctx.stroke()
  } else {
    const colors = ['#e81123', '#ffb900', '#107c41']
    ctx.fillStyle = colors[index] || '#107c41'
    ctx.beginPath()
    if (index === 0) {
      ctx.moveTo(x + r, y + r)
      ctx.lineTo(x, y - r * 0.5)
      ctx.lineTo(x + 2 * r, y - r * 0.5)
    } else if (index === 1) {
      ctx.moveTo(x + 2 * r, y)
      ctx.lineTo(x, y - r)
      ctx.lineTo(x, y + r)
    } else {
      ctx.moveTo(x + r, y - r)
      ctx.lineTo(x, y + r * 0.5)
      ctx.lineTo(x + 2 * r, y + r * 0.5)
    }
    ctx.closePath()
    ctx.fill()
  }
  ctx.restore()
}

/** Computes the layout rectangle of a cell relative to the grid viewport. */
export function getCellRect(
  storeInstance: typeof store,
  sheet: SheetState,
  r: number,
  c: number
): { x: number; y: number; w: number; h: number } | null {
  const z = sheet.layout.zoom / 100
  const showHead = sheet.layout.showHeadings
  const rowHeaderW = showHead ? 46 * z : 0
  const headerH = showHead ? HEADER_H * z : 0
  const colsAxis = buildAxis(sheet.layout.colWidths, sheet.layout.hiddenCols, DEFAULT_COL_WIDTH, MAX_COLS)
  const rowsAxis = buildAxis(sheet.layout.rowHeights, sheet.layout.hiddenRows, DEFAULT_ROW_HEIGHT, MAX_ROWS)
  const m = storeInstance.mergeAt(sheet, r, c) ?? { r1: r, c1: c, r2: r, c2: c }
  const x = rowHeaderW + colsAxis.offset(m.c1) * z
  const y = headerH + rowsAxis.offset(m.r1) * z
  const w = colsAxis.totalSize(m.c2 + 1) * z - colsAxis.totalSize(m.c1) * z
  const h = rowsAxis.totalSize(m.r2 + 1) * z - rowsAxis.totalSize(m.r1) * z
  return { x, y, w, h }
}

export default function Grid() {
  useStore()
  const ui = useUi()
  const sheet = store.active
  const scrollRef = useRef<HTMLDivElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const editorRef = useRef<HTMLTextAreaElement>(null)
  const hostRef = useRef<HTMLDivElement>(null)
  const [, setTick] = useState(0)
  const forceRedraw = useCallback(() => setTick((t) => t + 1), [])

  const [size, setSize] = useState({ w: 0, h: 0 })
  const [scroll, setScroll] = useState({ x: 0, y: 0 })
  const [extent, setExtent] = useState({ rows: 300, cols: 60 })
  const [ctxMenu, openCtxMenu, closeCtxMenu] = useMenu()
  const [ctxMenuKind, setCtxMenuKind] = useState<'cell' | 'col' | 'row'>('cell')
  const [resizeTip, setResizeTip] = useState<{ x: number; y: number; text: string } | null>(null)
  const [filterMenu, setFilterMenu] = useState<{ col: number; x: number; y: number } | null>(null)
  const [hoverCell, setHoverCell] = useState<CellRef | null>(null)
  const [listMenuPos, setListMenuPos] = useState<{ x: number; y: number } | null>(null)
  // Excel's Enter-after-Tab: while chaining Tab across cells in a single-cell selection, Enter
  // commits and moves down to the row below, but back to the column where the Tab-run started.
  const tabAnchorCol = useRef<number | null>(null)

  const drag = useRef<
    | null
    | { kind: 'cells'; anchor: CellRef; extendOnly: boolean; base: Range[] }
    | { kind: 'colHeader' | 'rowHeader'; anchorIdx: number; extendOnly: boolean; base: Range[] }
    | { kind: 'fill'; seed: Range }
    | { kind: 'resizeCol'; col: number; startX: number; startW: number; cols: number[] }
    | { kind: 'resizeRow'; row: number; startY: number; startH: number; rows: number[] }
  >(null)

  // ---------- layout metrics ----------
  const z = sheet.layout.zoom / 100
  const showHead = sheet.layout.showHeadings
  const showGrid = sheet.layout.showGridlines
  const used = store.usedRange(sheet)

  const colsAxis: Axis = useMemo(
    () => buildAxis(sheet.layout.colWidths, sheet.layout.hiddenCols, DEFAULT_COL_WIDTH, MAX_COLS),
    [sheet, sheet.layout.colWidths, sheet.layout.hiddenCols, ui]
  )
  const rowsAxis: Axis = useMemo(
    () => buildAxis(sheet.layout.rowHeights, sheet.layout.hiddenRows, DEFAULT_ROW_HEIGHT, MAX_ROWS),
    [sheet, sheet.layout.rowHeights, sheet.layout.hiddenRows, ui]
  )
  const rowDepth = useMemo(() => maxLevel(sheet.layout, 'row'), [sheet.layout.rowLevels])
  const colDepth = useMemo(() => maxLevel(sheet.layout, 'col'), [sheet.layout.colLevels])
  const { gutterL, gutterT } = gutterSizes(rowDepth, colDepth, showHead, ui.outlineHidden)
  const rowRuns = useMemo(() => outlineRuns(sheet.layout, 'row'), [sheet.layout.rowLevels, sheet.layout.hiddenRows, sheet.layout.outline])
  const colRuns = useMemo(() => outlineRuns(sheet.layout, 'col'), [sheet.layout.colLevels, sheet.layout.hiddenCols, sheet.layout.outline])
  const rowHeaderW = gutterL + (showHead ? Math.max(CORNER_W_MIN, 10 + String(Math.max(100, extent.rows)).length * 8) * z : 0)
  const headerH = gutterT + (showHead ? HEADER_H * z : 0)
  const frozen = sheet.layout.frozen
  const frozenW = colsAxis.totalSize(frozen.cols) * z
  const frozenH = rowsAxis.totalSize(frozen.rows) * z

  // reset scroll extent when sheet changes
  useEffect(() => {
    const r = Math.min(MAX_ROWS, Math.max(300, (used?.r2 ?? 0) + 100))
    const c = Math.min(MAX_COLS, Math.max(60, (used?.c2 ?? 0) + 40))
    setExtent({ rows: r, cols: c })
    if (scrollRef.current) { scrollRef.current.scrollTop = 0; scrollRef.current.scrollLeft = 0 }
    setScroll({ x: 0, y: 0 })
  }, [sheet.uid]) // eslint-disable-line react-hooks/exhaustive-deps

  const sizerW = colsAxis.totalSize(extent.cols) * z + 40
  const sizerH = rowsAxis.totalSize(extent.rows) * z + 40

  // ---------- resize observer ----------
  useLayoutEffect(() => {
    const el = hostRef.current
    if (!el) return
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }))
    ro.observe(el)
    setSize({ w: el.clientWidth, h: el.clientHeight })
    return () => ro.disconnect()
  }, [])

  const onScroll = useCallback(() => {
    const el = scrollRef.current
    if (!el) return
    setScroll({ x: el.scrollLeft, y: el.scrollTop })
    let grew = false
    const next = { ...extent }
    if (el.scrollLeft + el.clientWidth > sizerW - el.clientWidth && extent.cols < MAX_COLS) {
      next.cols = Math.min(MAX_COLS, extent.cols * 2)
      grew = true
    }
    if (el.scrollTop + el.clientHeight > sizerH - el.clientHeight && extent.rows < MAX_ROWS) {
      next.rows = Math.min(MAX_ROWS, extent.rows * 2)
      grew = true
    }
    if (grew) setExtent(next)
  }, [extent, sizerW, sizerH])

  // ---------- hit testing / geometry ----------
  function colAt(px: number): number {
    if (px < frozenW) return colsAxis.indexAt(px / z)
    return Math.max(frozen.cols, colsAxis.indexAt((px + scroll.x) / z))
  }
  function rowAt(py: number): number {
    if (py < frozenH) return rowsAxis.indexAt(py / z)
    return Math.max(frozen.rows, rowsAxis.indexAt((py + scroll.y) / z))
  }
  function cellScreenX(c: number): number {
    if (c < frozen.cols) return rowHeaderW + colsAxis.offset(c) * z
    return rowHeaderW + frozenW + colsAxis.offset(c) * z - colsAxis.offset(frozen.cols) * z - scroll.x
  }
  function cellScreenY(r: number): number {
    if (r < frozen.rows) return headerH + rowsAxis.offset(r) * z
    return headerH + frozenH + rowsAxis.offset(r) * z - rowsAxis.offset(frozen.rows) * z - scroll.y
  }
  function cellAt(x: number, y: number): CellRef {
    const cx = x - rowHeaderW, cy = y - headerH
    return { r: rowAt(Math.max(0, cy)), c: colAt(Math.max(0, cx)) }
  }
  function mergedRangeOf(r: number, c: number): Range {
    return store.mergeAt(sheet, r, c) ?? { r1: r, c1: c, r2: r, c2: c }
  }
  /** Screen rect of a cell (merge-expanded). */
  function cellRect(r: number, c: number): { x: number; y: number; w: number; h: number } {
    const m = mergedRangeOf(r, c)
    const x = cellScreenX(m.c1), y = cellScreenY(m.r1)
    const w = colsAxis.totalSize(m.c2 + 1) * z - colsAxis.totalSize(m.c1) * z
    const h = rowsAxis.totalSize(m.r2 + 1) * z - rowsAxis.totalSize(m.r1) * z
    return { x, y, w, h }
  }

  /** Column index whose AutoFilter dropdown arrow (r,c) = (autoFilter.range.r1, c) is at (x,y), if any. */
  function filterArrowAt(x: number, y: number): number | null {
    const af = sheet.layout.autoFilter
    if (!af) return null
    const r = af.range.r1
    const ry = cellScreenY(r)
    const rh = rowsAxis.size(r) * z
    if (y < ry || y > ry + rh) return null
    for (let c = af.range.c1; c <= af.range.c2; c++) {
      const cx = cellScreenX(c)
      const cw = colsAxis.size(c) * z
      if (cw <= 0) continue
      const boxW = Math.min(16, cw)
      if (x >= cx + cw - boxW && x <= cx + cw) return c
    }
    return null
  }

  function visibleCols(fromCol: number, untilCol: number, scrollPx: number, originX: number, availW: number): { c: number; x: number; w: number }[] {
    const out: { c: number; x: number; w: number }[] = []
    const base = colsAxis.offset(fromCol)
    let c = Math.max(fromCol, colsAxis.indexAt(base + scrollPx / z))
    let x = originX + ((colsAxis.offset(c) - base) * z - scrollPx)
    while (c < untilCol && x < originX + availW) {
      const w = colsAxis.size(c) * z
      if (w > 0.01) out.push({ c, x, w })
      x += w
      c++
    }
    return out
  }
  function visibleRows(fromRow: number, untilRow: number, scrollPx: number, originY: number, availH: number): { r: number; y: number; h: number }[] {
    const out: { r: number; y: number; h: number }[] = []
    // Offsets are relative to fromRow: the scrollable pane starts right below the frozen rows.
    const base = rowsAxis.offset(fromRow)
    let r = Math.max(fromRow, rowsAxis.indexAt(base + scrollPx / z))
    let y = originY + ((rowsAxis.offset(r) - base) * z - scrollPx)
    while (r < untilRow && y < originY + availH) {
      const h = rowsAxis.size(r) * z
      if (h > 0.01) out.push({ r, y, h })
      y += h
      r++
    }
    return out
  }

  // ---------- drawing ----------
  useLayoutEffect(() => {
    const canvas = canvasRef.current
    if (!canvas || size.w <= 0 || size.h <= 0) return
    const dpr = window.devicePixelRatio || 1
    const pxW = Math.round(size.w * dpr), pxH = Math.round(size.h * dpr)
    if (canvas.width !== pxW || canvas.height !== pxH) { canvas.width = pxW; canvas.height = pxH }
    const ctx = canvas.getContext('2d')!
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, size.w, size.h)
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, size.w, size.h)

    const sel = sheet.selection
    const primary = normRange(sel.ranges[sel.ranges.length - 1])
    // Cheap containment checks (never enumerate a whole-row/whole-column range's million+ indices).
    const normRanges = sel.ranges.map(normRange)
    const selectedCols = { has: (c: number) => normRanges.some((n) => c >= n.c1 && c <= n.c2) }
    const selectedRows = { has: (r: number) => normRanges.some((n) => r >= n.r1 && r <= n.r2) }
    const wholeColSel = primary.r1 === 0 && primary.r2 >= MAX_ROWS - 1
    const wholeRowSel = primary.c1 === 0 && primary.c2 >= MAX_COLS - 1

    function drawCellContent(r: number, c: number, x: number, y: number, w: number, h: number): void {
      if (w <= 0 || h <= 0) return
      const baseStyle = store.getStyle(sheet, r, c)
      const cf = evaluateCf(store, sheet, r, c)
      const style = cf.style ? { ...baseStyle, ...cf.style } : baseStyle
      const bg = cf.colorScale || style?.fill
      if (bg) { ctx.fillStyle = bg; ctx.fillRect(x, y, w, h) }

      if (cf.dataBar) {
        const barW = Math.max(0, Math.min(w - 2, Math.round((w - 2) * cf.dataBar.fraction)))
        if (barW > 0) {
          const grad = ctx.createLinearGradient(x + 1, y + 1, x + 1 + barW, y + 1)
          grad.addColorStop(0, cf.dataBar.color)
          grad.addColorStop(1, '#ffffff')
          ctx.fillStyle = grad
          ctx.fillRect(x + 1, y + 1, barW, h - 2)
          ctx.strokeStyle = cf.dataBar.color
          ctx.lineWidth = 1
          ctx.strokeRect(x + 1, y + 1, barW, h - 2)
        }
      }

      let iconPad = 0
      if (cf.icon) {
        iconPad = 14 * z
        const ix = x + 3 * z, iy = y + h / 2
        drawCfIcon(ctx, cf.icon.set, cf.icon.index, ix, iy, z)
      }

      const editingHere = ui.edit && ui.edit.r === r && ui.edit.c === c
      if (editingHere) return
      const disp = ui.formulasVisible ? { text: store.getRaw(sheet, r, c), color: undefined, align: 'left' as const } : store.getDisplay(sheet, r, c)
      if (!disp.text) return
      ctx.font = fontString(style, z)
      ctx.fillStyle = style?.color ?? disp.color ?? '#000000'
      const hAlign = style?.hAlign && style.hAlign !== 'general' ? style.hAlign : (disp.align ?? 'left')
      const vAlign = style?.vAlign ?? 'bottom'
      const pad = 3 * z
      ctx.textBaseline = 'middle'
      let tx = x + pad + iconPad
      ctx.textAlign = 'left'
      if (hAlign === 'right') { tx = x + w - pad; ctx.textAlign = 'right' }
      else if (hAlign === 'center' || hAlign === 'centerContinuous') { tx = x + w / 2; ctx.textAlign = 'center' }
      const ty = vAlign === 'top' ? y + pad + 6 * z : vAlign === 'middle' ? y + h / 2 : y + h - pad - 6 * z
      let text = disp.text
      const metrics = ctx.measureText(text)
      const isNumeric = typeof store.getValue(sheet, r, c) === 'number'
      if (metrics.width > w - 2 * pad - iconPad && isNumeric) {
        const avail = w - 2 * pad - iconPad
        const fmt = style?.numFmt
        // General numbers shrink the way Excel's do (fewer decimals, then 2.63E+09) before ####.
        const fit = !fmt || fmt === 'General' ? fitGeneral(store.getValue(sheet, r, c) as number, (t) => ctx.measureText(t).width <= avail) : null
        text = fit ?? '#'.repeat(Math.max(1, Math.floor(avail / (7 * z))))
      } else if (metrics.width > w - 2 * pad - iconPad && !style?.wrap) {
        // overflow into empty neighbours: only clip if neighbour is non-empty
      }
      ctx.save()
      ctx.beginPath()
      // allow overflow to the right/left for text into empty neighbour cells: clip to a wider box
      let clipX = x, clipW = w
      if (!style?.wrap && hAlign === 'left') {
        let cc = c + 1
        while (clipW < metrics.width + 2 * pad + iconPad && cc < MAX_COLS && store.getRaw(sheet, r, cc) === '' && !store.mergeAt(sheet, r, cc)) {
          clipW += colsAxis.size(cc) * z
          cc++
        }
      } else if (!style?.wrap && hAlign === 'right') {
        let cc = c - 1
        while (clipW < metrics.width + 2 * pad + iconPad && cc >= 0 && store.getRaw(sheet, r, cc) === '' && !store.mergeAt(sheet, r, cc)) {
          clipX -= colsAxis.size(cc) * z
          clipW += colsAxis.size(cc) * z
          cc--
        }
      }
      ctx.rect(clipX, y, clipW, h)
      ctx.clip()
      ctx.fillText(text, tx, ty)
      ctx.restore()
      if (style?.underline) {
        const uw = metrics.width
        const ux = hAlign === 'right' ? tx - uw : hAlign === 'center' ? tx - uw / 2 : tx
        ctx.strokeStyle = ctx.fillStyle as string
        ctx.beginPath(); ctx.moveTo(ux, ty + 5 * z); ctx.lineTo(ux + uw, ty + 5 * z); ctx.stroke()
      }
      if (style?.strike) {
        const uw = metrics.width
        const ux = hAlign === 'right' ? tx - uw : hAlign === 'center' ? tx - uw / 2 : tx
        ctx.strokeStyle = ctx.fillStyle as string
        ctx.beginPath(); ctx.moveTo(ux, ty); ctx.lineTo(ux + uw, ty); ctx.stroke()
      }
    }

    function drawBorders(r: number, c: number, x: number, y: number, w: number, h: number): void {
      const baseStyle = store.getStyle(sheet, r, c)
      const cf = evaluateCf(store, sheet, r, c)
      const style = cf.style?.borders ? { ...baseStyle, borders: { ...baseStyle?.borders, ...cf.style.borders } } : baseStyle
      const b = style?.borders
      if (!b) return
      ctx.lineWidth = 1
      const draw = (side: 'top' | 'right' | 'bottom' | 'left') => {
        const s = b[side]
        if (!s) return
        ctx.strokeStyle = s.color ?? '#000000'
        ctx.lineWidth = s.style === 'thick' || s.style === 'medium' ? 2 : 1
        ctx.setLineDash(s.style === 'dashed' || s.style === 'mediumDashed' ? [3, 2] : s.style === 'dotted' || s.style === 'hair' ? [1, 1] : [])
        ctx.beginPath()
        if (side === 'top') { ctx.moveTo(x, y + 0.5); ctx.lineTo(x + w, y + 0.5) }
        if (side === 'bottom') { ctx.moveTo(x, y + h - 0.5); ctx.lineTo(x + w, y + h - 0.5) }
        if (side === 'left') { ctx.moveTo(x + 0.5, y); ctx.lineTo(x + 0.5, y + h) }
        if (side === 'right') { ctx.moveTo(x + w - 0.5, y); ctx.lineTo(x + w - 0.5, y + h) }
        ctx.stroke()
        ctx.setLineDash([])
      }
      draw('top'); draw('right'); draw('bottom'); draw('left')
    }

    function drawCellsRegion(fromR: number, untilR: number, fromC: number, untilC: number, sx: number, sy: number, ox: number, oy: number, aw: number, ah: number): void {
      const rows = visibleRows(fromR, untilR, sy, oy, ah)
      const cols = visibleCols(fromC, untilC, sx, ox, aw)
      ctx.save()
      ctx.beginPath(); ctx.rect(ox, oy, aw, ah); ctx.clip()
      for (const { r, y, h } of rows) for (const { c, x, w } of cols) {
        const m = store.mergeAt(sheet, r, c)
        if (m && (m.r1 !== r || m.c1 !== c)) continue // drawn from its top-left
        if (m) { const rc = cellRect(r, c); drawCellContent(r, c, rc.x, rc.y, rc.w, rc.h) }
        else drawCellContent(r, c, x, y, w, h)
      }
      if (showGrid) {
        ctx.strokeStyle = '#e1e1e1'
        ctx.lineWidth = 1
        ctx.beginPath()
        for (const { y } of rows) { ctx.moveTo(ox, Math.round(y) + 0.5); ctx.lineTo(ox + aw, Math.round(y) + 0.5) }
        const lastRow = rows[rows.length - 1]
        if (lastRow) { const yy = Math.round(lastRow.y + lastRow.h) + 0.5; ctx.moveTo(ox, yy); ctx.lineTo(ox + aw, yy) }
        for (const { x } of cols) { ctx.moveTo(Math.round(x) + 0.5, oy); ctx.lineTo(Math.round(x) + 0.5, oy + ah) }
        const lastCol = cols[cols.length - 1]
        if (lastCol) { const xx = Math.round(lastCol.x + lastCol.w) + 0.5; ctx.moveTo(xx, oy); ctx.lineTo(xx, oy + ah) }
        ctx.stroke()
      }
      // borders + merge outlines drawn after gridlines
      for (const { r, y, h } of rows) for (const { c, x, w } of cols) {
        const m = store.mergeAt(sheet, r, c)
        if (m) { if (m.r1 === r && m.c1 === c) { const rc = cellRect(r, c); drawBorders(r, c, rc.x, rc.y, rc.w, rc.h) } }
        else drawBorders(r, c, x, y, w, h)
      }
      ctx.restore()
    }

    // main scrollable body (BR)
    drawCellsRegion(frozen.rows, extent.rows, frozen.cols, extent.cols, scroll.x, scroll.y, rowHeaderW + frozenW, headerH + frozenH, size.w - rowHeaderW - frozenW, size.h - headerH - frozenH)
    // TR: frozen rows, scrollable cols
    if (frozen.rows > 0) drawCellsRegion(0, frozen.rows, frozen.cols, extent.cols, scroll.x, 0, rowHeaderW + frozenW, headerH, size.w - rowHeaderW - frozenW, frozenH)
    // BL: scrollable rows, frozen cols
    if (frozen.cols > 0) drawCellsRegion(frozen.rows, extent.rows, 0, frozen.cols, 0, scroll.y, rowHeaderW, headerH + frozenH, frozenW, size.h - headerH - frozenH)
    // TL: frozen both
    if (frozen.rows > 0 && frozen.cols > 0) drawCellsRegion(0, frozen.rows, 0, frozen.cols, 0, 0, rowHeaderW, headerH, frozenW, frozenH)

    // frozen divider lines
    if (frozen.rows > 0) { ctx.strokeStyle = '#616161'; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(rowHeaderW, headerH + frozenH); ctx.lineTo(size.w, headerH + frozenH); ctx.stroke() }
    if (frozen.cols > 0) { ctx.strokeStyle = '#616161'; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(rowHeaderW + frozenW, headerH); ctx.lineTo(rowHeaderW + frozenW, size.h); ctx.stroke() }

    // ---------- selection overlay ----------
    function forEachVisibleRectOf(rg: Range, fn: (x: number, y: number, w: number, h: number) => void): void {
      const n = normRange(rg)
      const c1 = Math.max(n.c1, 0), c2 = Math.min(n.c2, extent.cols - 1)
      const r1 = Math.max(n.r1, 0), r2 = Math.min(n.r2, extent.rows - 1)
      if (c1 > c2 || r1 > r2) return
      const x1 = cellScreenX(c1), x2 = cellScreenX(Math.min(c2, extent.cols - 1)) + colsAxis.size(c2) * z
      const y1 = cellScreenY(r1), y2 = cellScreenY(Math.min(r2, extent.rows - 1)) + rowsAxis.size(r2) * z
      fn(Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1))
    }

    ctx.save()
    ctx.beginPath(); ctx.rect(rowHeaderW, headerH, size.w - rowHeaderW, size.h - headerH); ctx.clip()
    sel.ranges.forEach((rg, i) => {
      const isPrimary = i === sel.ranges.length - 1
      forEachVisibleRectOf(rg, (x, y, w, h) => {
        ctx.fillStyle = 'rgba(33,115,70,0.08)'
        ctx.fillRect(x, y, w, h)
        ctx.strokeStyle = '#107c41'
        ctx.lineWidth = isPrimary ? 2 : 1.5
        ctx.strokeRect(x + ctx.lineWidth / 2, y + ctx.lineWidth / 2, Math.max(0, w - ctx.lineWidth), Math.max(0, h - ctx.lineWidth))
      })
    })
    // clear active-cell fill (Excel shows the active cell un-tinted) + fill handle. Re-draw the
    // cell's own content afterwards: clearRect/fillRect wipe whatever was painted underneath
    // (including the cell's text), so without this the active cell would render blank.
    if (!wholeColSel && !wholeRowSel) {
      const ar = mergedRangeOf(sel.active.r, sel.active.c)
      forEachVisibleRectOf(ar, (x, y, w, h) => {
        ctx.clearRect(x + 2, y + 2, Math.max(0, w - 4), Math.max(0, h - 4))
        ctx.fillStyle = '#ffffff'
        ctx.fillRect(x + 2, y + 2, Math.max(0, w - 4), Math.max(0, h - 4))
        drawCellContent(ar.r1, ar.c1, x, y, w, h)
      })
    }
    // fill handle at bottom-right of primary range
    forEachVisibleRectOf(primary, (x, y, w, h) => {
      const hs = 6
      ctx.fillStyle = '#107c41'
      ctx.fillRect(x + w - hs / 2, y + h - hs / 2, hs, hs)
      ctx.strokeStyle = '#ffffff'; ctx.lineWidth = 1
      ctx.strokeRect(x + w - hs / 2, y + h - hs / 2, hs, hs)
    })
    // formula reference highlight boxes
    if (ui.edit && ui.edit.text.startsWith('=')) {
      const refs = extractRefColors(ui.edit.text)
      refs.forEach(({ range, color }) => {
        forEachVisibleRectOf(range, (x, y, w, h) => {
          ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.setLineDash([])
          ctx.strokeRect(x + 1, y + 1, Math.max(0, w - 2), Math.max(0, h - 2))
        })
      })
    }
    ctx.restore()

    // ---------- AutoFilter dropdown arrows ----------
    const af = sheet.layout.autoFilter
    if (af) {
      const hr = af.range.r1
      const ry = cellScreenY(hr)
      const rh = rowsAxis.size(hr) * z
      if (rh > 0 && ry < size.h && ry + rh > headerH) {
        ctx.save()
        ctx.beginPath(); ctx.rect(rowHeaderW, headerH, size.w - rowHeaderW, size.h - headerH); ctx.clip()
        for (let c = af.range.c1; c <= af.range.c2; c++) {
          const cx = cellScreenX(c)
          const cw = colsAxis.size(c) * z
          if (cw <= 0 || cx > size.w || cx + cw < rowHeaderW) continue
          const boxW = Math.min(16, cw)
          const bx = cx + cw - boxW
          ctx.fillStyle = '#eaeaea'
          ctx.fillRect(bx, ry + 1, boxW - 1, rh - 2)
          ctx.strokeStyle = '#adadad'
          ctx.strokeRect(bx + 0.5, ry + 1.5, boxW - 2, rh - 3)
          const filtered = !!af.columns[c]
          ctx.fillStyle = filtered ? '#107c41' : '#616161'
          const mx = bx + boxW / 2, my = ry + rh / 2
          ctx.beginPath()
          ctx.moveTo(mx - 4, my - 2); ctx.lineTo(mx + 4, my - 2); ctx.lineTo(mx, my + 3)
          ctx.closePath(); ctx.fill()
        }
        ctx.restore()
      }
    }

    // ---------- note & comment indicators ----------
    if (Object.keys(sheet.layout.notes).length || Object.keys(sheet.layout.comments).length) {
      ctx.save()
      ctx.beginPath(); ctx.rect(rowHeaderW, headerH, size.w - rowHeaderW, size.h - headerH); ctx.clip()
      const drawCorner = (key: string, color: string) => {
        const [rr, cc] = key.split(',').map(Number)
        const rc = cellRect(rr, cc)
        if (rc.x + rc.w < rowHeaderW || rc.x > size.w || rc.y + rc.h < headerH || rc.y > size.h) return
        ctx.fillStyle = color
        ctx.beginPath()
        ctx.moveTo(rc.x + rc.w - 6, rc.y)
        ctx.lineTo(rc.x + rc.w, rc.y)
        ctx.lineTo(rc.x + rc.w, rc.y + 6)
        ctx.closePath()
        ctx.fill()
      }
      Object.keys(sheet.layout.notes).forEach((k) => drawCorner(k, '#c00000'))
      Object.keys(sheet.layout.comments).forEach((k) => drawCorner(k, '#7719aa'))
      ctx.restore()
    }

    // ---------- headers ----------
    ctx.save()
    ctx.beginPath(); ctx.rect(0, 0, size.w, size.h); ctx.clip()
    if (showHead) {
      // corner
      ctx.fillStyle = '#f5f5f5'
      ctx.fillRect(gutterL, gutterT, rowHeaderW - gutterL, headerH - gutterT)
      ctx.strokeStyle = '#d4d4d4'; ctx.strokeRect(gutterL + 0.5, gutterT + 0.5, rowHeaderW - gutterL - 1, headerH - gutterT - 1)
      ctx.beginPath()
      ctx.fillStyle = '#616161'
      ctx.moveTo(rowHeaderW - 8, headerH - 4); ctx.lineTo(rowHeaderW - 3, headerH - 4); ctx.lineTo(rowHeaderW - 3, headerH - 9)
      ctx.fill()

      const drawColHeaders = (fromC: number, untilC: number, sx: number, ox: number, aw: number) => {
        const cols = visibleCols(fromC, untilC, sx, ox, aw)
        for (const { c, x, w } of cols) {
          const isSel = selectedCols.has(c) && !wholeColSel
          ctx.fillStyle = isSel ? '#d3f0e0' : '#f5f5f5'
          ctx.fillRect(x, gutterT, w, headerH - gutterT)
          ctx.strokeStyle = '#d4d4d4'
          ctx.beginPath(); ctx.moveTo(x + w - 0.5, gutterT); ctx.lineTo(x + w - 0.5, headerH); ctx.stroke()
          ctx.font = `${isSel ? '600 ' : ''}${Math.round(11 * z)}px "Segoe UI", sans-serif`
          ctx.fillStyle = isSel ? '#107c41' : '#616161'
          ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
          ctx.fillText(colName(c), x + w / 2, gutterT + (headerH - gutterT) / 2 + 1)
          if (isSel) { ctx.strokeStyle = '#107c41'; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(x, headerH - 1); ctx.lineTo(x + w, headerH - 1); ctx.stroke() }
        }
      }
      drawColHeaders(frozen.cols, extent.cols, scroll.x, rowHeaderW + frozenW, size.w - rowHeaderW - frozenW)
      if (frozen.cols > 0) drawColHeaders(0, frozen.cols, 0, rowHeaderW, frozenW)

      const filterActive = !!af && af.filteredOutRows.length > 0
      const drawRowHeaders = (fromR: number, untilR: number, sy: number, oy: number, ah: number) => {
        const rows = visibleRows(fromR, untilR, sy, oy, ah)
        for (const { r, y, h } of rows) {
          const isSel = selectedRows.has(r) && !wholeRowSel
          const isFilteredRow = filterActive && af!.range && r > af!.range.r1 && r <= af!.range.r2 && !sheet.layout.hiddenRows[r]
          ctx.fillStyle = isSel ? '#d3f0e0' : '#f5f5f5'
          ctx.fillRect(gutterL, y, rowHeaderW - gutterL, h)
          ctx.strokeStyle = '#d4d4d4'
          ctx.beginPath(); ctx.moveTo(gutterL, y + h - 0.5); ctx.lineTo(rowHeaderW, y + h - 0.5); ctx.stroke()
          ctx.font = `${isSel ? '600 ' : ''}${Math.round(11 * z)}px "Segoe UI", sans-serif`
          ctx.fillStyle = isSel ? '#107c41' : isFilteredRow ? '#1a73c9' : '#616161'
          ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
          ctx.fillText(String(r + 1), gutterL + (rowHeaderW - gutterL) / 2, y + h / 2 + 1)
          if (isSel) { ctx.strokeStyle = '#107c41'; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(rowHeaderW - 1, y); ctx.lineTo(rowHeaderW - 1, y + h); ctx.stroke() }
        }
      }
      drawRowHeaders(frozen.rows, extent.rows, scroll.y, headerH + frozenH, size.h - headerH - frozenH)
      if (frozen.rows > 0) drawRowHeaders(0, frozen.rows, 0, headerH, frozenH)
    }
    drawGutter(ctx, gutterGeo())
    ctx.restore()
  })

  // ---------- extract colored refs from a formula for Point-mode highlighting ----------
  function extractRefColors(text: string): { range: Range; color: string }[] {
    const re = /(?:^|[^A-Za-z0-9_$])(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?/g
    const out: { range: Range; color: string }[] = []
    let m: RegExpExecArray | null
    let i = 0
    while ((m = re.exec(text))) {
      const rg1 = parseA1(m[1])
      if (!rg1) continue
      const rg2 = m[2] ? parseA1(m[2]) : null
      const range = rg2 ? normRange({ r1: rg1.r1, c1: rg1.c1, r2: rg2.r2, c2: rg2.c2 }) : rg1
      out.push({ range, color: REF_COLORS[i % REF_COLORS.length] })
      i++
    }
    return out
  }

  // ---------- editor positioning ----------
  useLayoutEffect(() => {
    const e = ui.edit
    const ta = editorRef.current
    if (!e || !ta) return
    if (e.source !== 'cell') return
    const rc = cellRect(e.r, e.c)
    const font = fontString(store.getStyle(sheet, e.r, e.c), z)
    ta.style.font = font
    const ctx = canvasRef.current?.getContext('2d')
    let measured = rc.w
    if (ctx) {
      ctx.font = font
      const lines = e.text.split('\n')
      const widest = Math.max(...lines.map((l) => ctx.measureText(l || ' ').width))
      measured = widest + 16
    }
    ta.style.left = rc.x + 'px'
    ta.style.top = rc.y + 'px'
    ta.style.width = Math.max(rc.w, measured) + 'px'
    ta.style.height = Math.max(rc.h, 18) + 'px'
    // Keep the DOM caret in sync with edit.caret: on a fresh edit (esp. the first keystroke that
    // starts "Enter" mode) the freshly-focused textarea otherwise leaves the native caret at 0,
    // so the next keystrokes land BEFORE the first character instead of after it.
    const caret = e.caret ?? e.text.length
    if (ta.selectionStart !== caret || ta.selectionEnd !== caret) {
      try { ta.setSelectionRange(caret, caret) } catch { /* ignore */ }
    }
  })

  // ---------- selection setting ----------
  function setSel(ranges: Range[], active: CellRef): void {
    store.setSelection({ ranges, active }, sheet)
  }
  function selectSingle(r: number, c: number): void {
    const m = mergedRangeOf(r, c)
    setSel([{ r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }], { r: m.r1, c: m.c1 })
  }

  function beginEdit(r: number, c: number, mode: 'Enter' | 'Edit', initialText?: string): void {
    if (!store.canEditCells(sheet, [{ r, c }])) return
    pointCur.current = null
    pointedLen.current = 0
    const raw = initialText ?? store.getRaw(sheet, r, c)
    setUi({ edit: { r, c, text: initialText ?? (mode === 'Enter' ? '' : raw), caret: (initialText ?? raw).length, source: 'cell', mode } })
    store.setMode(mode)
    if (mode === 'Enter' && initialText === undefined) {
      setUi({ edit: { r, c, text: '', caret: 0, source: 'cell', mode: 'Enter' } })
    }
  }

  function moveActive(dr: number, dc: number, extend: boolean, opts?: { jump?: boolean }): void {
    const cur = sheet.selection
    // Shift+Arrow moves the far corner of the selection (Excel), keeping the active cell as anchor.
    const last = cur.ranges[cur.ranges.length - 1]
    const from: CellRef = extend && last
      ? { r: last.r1 === cur.active.r ? last.r2 : last.r1, c: last.c1 === cur.active.c ? last.c2 : last.c1 }
      : cur.active
    let r: number, c: number
    if (opts?.jump) {
      const j = jumpTo(from.r, from.c, dr, dc)
      r = j.r; c = j.c
    } else if (extend) {
      r = from.r + dr
      c = from.c + dc
    } else {
      const m = mergedRangeOf(cur.active.r, cur.active.c)
      r = dr < 0 ? m.r1 + dr : dr > 0 ? m.r2 + dr : cur.active.r
      c = dc < 0 ? m.c1 + dc : dc > 0 ? m.c2 + dc : cur.active.c
    }
    r = Math.max(0, Math.min(MAX_ROWS - 1, r))
    c = Math.max(0, Math.min(MAX_COLS - 1, c))
    if (extend) {
      const ranges = [...cur.ranges]
      ranges[ranges.length - 1] = { r1: cur.active.r, c1: cur.active.c, r2: r, c2: c }
      setSel(ranges, cur.active)
    } else {
      selectSingle(r, c)
    }
    ensureVisible(r, c)
  }
  function r0(a: CellRef) { return a.r }
  function c0(a: CellRef) { return a.c }

  function jumpTo(r: number, c: number, dr: number, dc: number): CellRef {
    const hasVal = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS && store.getRaw(sheet, rr, cc) !== ''
    let rr = r, cc = c
    const step = () => { rr += dr; cc += dc }
    const inBounds = () => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS
    if (!hasVal(rr, cc)) {
      // move until hitting a value or edge
      step()
      while (inBounds() && !hasVal(rr, cc)) step()
      if (!inBounds()) { rr = Math.max(0, Math.min(MAX_ROWS - 1, rr)); cc = Math.max(0, Math.min(MAX_COLS - 1, cc)); return { r: rr, c: cc } }
      return { r: rr, c: cc }
    }
    // currently on a value: move while next has value too
    const startHas = hasVal(rr + dr, cc + dc)
    if (startHas) {
      while (inBounds() && hasVal(rr + dr, cc + dc)) step()
    } else {
      while (inBounds() && hasVal(rr, cc)) step()
      // back off one since loop stepped onto empty
      rr -= dr; cc -= dc
      step()
      while (inBounds() && !hasVal(rr, cc)) step()
    }
    rr = Math.max(0, Math.min(MAX_ROWS - 1, rr))
    cc = Math.max(0, Math.min(MAX_COLS - 1, cc))
    return { r: rr, c: cc }
  }

  function ensureVisible(r: number, c: number): void {
    const el = scrollRef.current
    if (!el) return
    const x0 = (colsAxis.offset(c) - colsAxis.offset(frozen.cols)) * z
    const x1 = x0 + colsAxis.size(c) * z
    const avail = size.w - rowHeaderW - frozenW
    if (r < frozen.rows === false) {
      if (x0 < el.scrollLeft) el.scrollLeft = x0
      else if (x1 > el.scrollLeft + avail) el.scrollLeft = x1 - avail
    }
    const y0 = (rowsAxis.offset(r) - rowsAxis.offset(frozen.rows)) * z
    const y1 = y0 + rowsAxis.size(r) * z
    const availH = size.h - headerH - frozenH
    if (y0 < el.scrollTop) el.scrollTop = y0
    else if (y1 > el.scrollTop + availH) el.scrollTop = y1 - availH
    onScroll()
  }

  // ---------- mouse ----------
  function inFillHandle(x: number, y: number): boolean {
    const sel = sheet.selection
    const primary = normRange(sel.ranges[sel.ranges.length - 1])
    let hit = false
    const c2 = Math.min(primary.c2, extent.cols - 1), r2 = Math.min(primary.r2, extent.rows - 1)
    const fx = cellScreenX(c2) + colsAxis.size(c2) * z
    const fy = cellScreenY(r2) + rowsAxis.size(r2) * z
    if (Math.abs(x - fx) < 5 && Math.abs(y - fy) < 5) hit = true
    return hit
  }

  function findColBorder(x: number, y: number): number | null {
    if (y > headerH) return null
    for (const c of [...Array(60).keys()]) {
      const col = colAt(x) - 2 + c
      if (col < 0) continue
      const right = cellScreenX(col) + colsAxis.size(col) * z
      if (Math.abs(x - right) < 4) return col
    }
    return null
  }
  function findRowBorder(x: number, y: number): number | null {
    if (x > rowHeaderW) return null
    for (const r of [...Array(60).keys()]) {
      const row = rowAt(y) - 2 + r
      if (row < 0) continue
      const bottom = cellScreenY(row) + rowsAxis.size(row) * z
      if (Math.abs(y - bottom) < 4) return row
    }
    return null
  }

  function gutterGeo(): GutterGeo {
    return {
      gutterL, gutterT, headerH, rowHeaderW, showHead, size, rowRuns, colRuns, rowDepth, colDepth,
      rowY: (i) => cellScreenY(i), colX: (i) => cellScreenX(i),
      summaryBelow: sheet.layout.outline?.summaryBelow !== false, summaryRight: sheet.layout.outline?.summaryRight !== false
    }
  }

  const linkClick = useRef<{ r: number; c: number; t: number } | null>(null)
  function onMouseDown(e: React.MouseEvent): void {
    if (e.button === 2) return
    tabAnchorCol.current = null
    // A chart's own mousedown handler stops propagation, so getting here means the click landed
    // on the grid itself: deselect any floating chart.
    if (getUi().selectedChart) setUi({ selectedChart: null })
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    // Point mode: editing a formula where a reference can go (after =, an operator, "(" or ","),
    // or right after a pointed reference: clicking inserts/replaces a reference instead of committing.
    if (canPoint() && y > headerH && x > rowHeaderW) {
      e.preventDefault() // keep the editor focused
      const { r, c } = cellAt(x, y)
      pointCur.current = { ar: r, ac: c, r, c }
      insertRefIntoEdit(r, c, r, c)
      requestAnimationFrame(() => editorRef.current?.focus())
      drag.current = { kind: 'cells', anchor: { r, c }, extendOnly: true, base: [] }
      window.addEventListener('mousemove', onPointDrag)
      window.addEventListener('mouseup', onPointUp)
      return
    }
    editorRef.current?.blur()

    // outline gutter: level buttons and [-]/[+]; the rest of the gutter is inert
    if (gutterL > 0 || gutterT > 0) {
      const hit = hitGutter(gutterGeo(), x, y)
      if (hit) { if (hit.kind === 'level') showOutlineLevel(hit.axis, hit.n); else toggleRun(hit.axis, hit.run); return }
      if (x < gutterL || y < gutterT) return
    }

    // select-all corner
    if (showHead && x < rowHeaderW && y < headerH) { cmd.selectAll(); return }

    // column resize
    const rcol = findColBorder(x, y)
    if (rcol !== null && y < headerH) {
      drag.current = { kind: 'resizeCol', col: rcol, startX: x, startW: colsAxis.size(rcol), cols: selectedColsContains(rcol) ? [...selectedColsSet()] : [rcol] }
      window.addEventListener('mousemove', onResizeMove)
      window.addEventListener('mouseup', onResizeUp)
      return
    }
    const rrow = findRowBorder(x, y)
    if (rrow !== null && x < rowHeaderW) {
      drag.current = { kind: 'resizeRow', row: rrow, startY: y, startH: rowsAxis.size(rrow), rows: selectedRowsContains(rrow) ? [...selectedRowsSet()] : [rrow] }
      window.addEventListener('mousemove', onResizeMove)
      window.addEventListener('mouseup', onResizeUp)
      return
    }

    // fill handle
    if (getUi().edit === null && inFillHandle(x, y)) {
      const sel = sheet.selection
      const seed = normRange(sel.ranges[sel.ranges.length - 1])
      drag.current = { kind: 'fill', seed }
      window.addEventListener('mousemove', onFillDrag)
      window.addEventListener('mouseup', onFillUp)
      return
    }

    // AutoFilter dropdown arrow
    if (x >= rowHeaderW && y >= headerH) {
      const fc = filterArrowAt(x, y)
      if (fc !== null) {
        if (getUi().edit) commitEdit()
        const af = sheet.layout.autoFilter!
        const hostRect = hostRef.current!.getBoundingClientRect() // menu is position:fixed, so use client coordinates
        setFilterMenu({ col: fc, x: hostRect.left + cellScreenX(fc), y: hostRect.top + cellScreenY(af.range.r1) + rowsAxis.size(af.range.r1) * z })
        return
      }
    }

    // committed pending edit if clicking elsewhere without point mode
    if (getUi().edit) commitEdit()

    // column header
    if (showHead && y < headerH && x >= rowHeaderW) {
      const c = colAt(x)
      const extend = e.shiftKey
      const addRange = e.ctrlKey || e.metaKey
      if (extend) {
        const cur = sheet.selection
        const ranges = [...cur.ranges]
        ranges[ranges.length - 1] = { r1: 0, c1: cur.active.c, r2: MAX_ROWS - 1, c2: c }
        setSel(ranges, cur.active)
      } else {
        const rg: Range = { r1: 0, c1: c, r2: MAX_ROWS - 1, c2: c }
        setSel(addRange ? [...sheet.selection.ranges, rg] : [rg], { r: 0, c })
      }
      drag.current = { kind: 'colHeader', anchorIdx: c, extendOnly: false, base: addRange ? sheet.selection.ranges.slice(0, -1) : [] }
      window.addEventListener('mousemove', onHeaderDrag)
      window.addEventListener('mouseup', onHeaderUp)
      return
    }
    // row header
    if (showHead && x < rowHeaderW && y >= headerH) {
      const r = rowAt(y)
      const extend = e.shiftKey
      const addRange = e.ctrlKey || e.metaKey
      if (extend) {
        const cur = sheet.selection
        const ranges = [...cur.ranges]
        ranges[ranges.length - 1] = { r1: cur.active.r, c1: 0, r2: r, c2: MAX_COLS - 1 }
        setSel(ranges, cur.active)
      } else {
        const rg: Range = { r1: r, c1: 0, r2: r, c2: MAX_COLS - 1 }
        setSel(addRange ? [...sheet.selection.ranges, rg] : [rg], { r, c: 0 })
      }
      drag.current = { kind: 'rowHeader', anchorIdx: r, extendOnly: false, base: addRange ? sheet.selection.ranges.slice(0, -1) : [] }
      window.addEventListener('mousemove', onHeaderDrag)
      window.addEventListener('mouseup', onHeaderUp)
      return
    }

    if (y < headerH || x < rowHeaderW) return
    const { r, c } = cellAt(x, y)
    const extend = e.shiftKey
    const addRange = e.ctrlKey || e.metaKey
    if (extend) {
      const cur = sheet.selection
      const ranges = [...cur.ranges]
      ranges[ranges.length - 1] = { r1: cur.active.r, c1: cur.active.c, r2: r, c2: c }
      setSel(ranges, cur.active)
    } else if (addRange) {
      setSel([...sheet.selection.ranges, { r1: r, c1: c, r2: r, c2: c }], { r, c })
    } else {
      selectSingle(r, c)
      // Excel follows a hyperlink on a quick click; click-and-hold just selects the cell.
      linkClick.current = linkAt(sheet, r, c) ? { r, c, t: Date.now() } : null
    }
    drag.current = { kind: 'cells', anchor: { r, c }, extendOnly: false, base: [] }
    window.addEventListener('mousemove', onCellDrag)
    window.addEventListener('mouseup', onCellUp)
  }

  function selectedColsSet(): Set<number> {
    const s = new Set<number>()
    sheet.selection.ranges.forEach((rg) => { const n = normRange(rg); for (let c = n.c1; c <= n.c2; c++) s.add(c) })
    return s
  }
  function selectedRowsSet(): Set<number> {
    const s = new Set<number>()
    sheet.selection.ranges.forEach((rg) => { const n = normRange(rg); for (let r = n.r1; r <= n.r2; r++) s.add(r) })
    return s
  }
  function selectedColsContains(c: number): boolean { return selectedColsSet().has(c) }
  function selectedRowsContains(r: number): boolean { return selectedRowsSet().has(r) }

  function autoscrollIfNeeded(x: number, y: number): void {
    const el = scrollRef.current
    if (!el) return
    const edge = 24
    if (x > size.w - edge) el.scrollLeft += 20
    else if (x < rowHeaderW + edge && x > rowHeaderW) el.scrollLeft -= 20
    if (y > size.h - edge) el.scrollTop += 20
    else if (y < headerH + edge && y > headerH) el.scrollTop -= 20
  }

  function onCellDrag(e: MouseEvent): void {
    if (!drag.current || drag.current.kind !== 'cells') return
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    autoscrollIfNeeded(x, y)
    const { r, c } = cellAt(Math.max(rowHeaderW, x), Math.max(headerH, y))
    const cur = sheet.selection
    const ranges = [...cur.ranges]
    ranges[ranges.length - 1] = { r1: drag.current.anchor.r, c1: drag.current.anchor.c, r2: r, c2: c }
    setSel(ranges, drag.current.anchor)
  }
  function onCellUp(): void {
    const lc = linkClick.current
    linkClick.current = null
    if (lc && Date.now() - lc.t < 700) {
      const n = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
      const link = linkAt(sheet, lc.r, lc.c)
      if (link && n.r1 === lc.r && n.r2 === lc.r && n.c1 === lc.c && n.c2 === lc.c) void followLink(link)
    }
    drag.current = null
    window.removeEventListener('mousemove', onCellDrag)
    window.removeEventListener('mouseup', onCellUp)
  }

  function onPointDrag(e: MouseEvent): void {
    if (!drag.current || drag.current.kind !== 'cells') return
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    const { r, c } = cellAt(Math.max(rowHeaderW, x), Math.max(headerH, y))
    insertRefIntoEdit(drag.current.anchor.r, drag.current.anchor.c, r, c)
  }
  function onPointUp(): void {
    drag.current = null
    window.removeEventListener('mousemove', onPointDrag)
    window.removeEventListener('mouseup', onPointUp)
  }

  function onHeaderDrag(e: MouseEvent): void {
    if (!drag.current || (drag.current.kind !== 'colHeader' && drag.current.kind !== 'rowHeader')) return
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    autoscrollIfNeeded(x, y)
    const cur = sheet.selection
    if (drag.current.kind === 'colHeader') {
      const c = colAt(Math.max(rowHeaderW, x))
      const rg: Range = { r1: 0, c1: drag.current.anchorIdx, r2: MAX_ROWS - 1, c2: c }
      setSel([...drag.current.base, rg], { r: 0, c: drag.current.anchorIdx })
    } else {
      const r = rowAt(Math.max(headerH, y))
      const rg: Range = { r1: drag.current.anchorIdx, c1: 0, r2: r, c2: MAX_COLS - 1 }
      setSel([...drag.current.base, rg], { r: drag.current.anchorIdx, c: 0 })
    }
    void cur
  }
  function onHeaderUp(): void {
    drag.current = null
    window.removeEventListener('mousemove', onHeaderDrag)
    window.removeEventListener('mouseup', onHeaderUp)
  }

  function onFillDrag(e: MouseEvent): void {
    if (!drag.current || drag.current.kind !== 'fill') return
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    autoscrollIfNeeded(x, y)
    const { r, c } = cellAt(Math.max(rowHeaderW, x), Math.max(headerH, y))
    const seed = drag.current.seed
    const target = normRange({
      r1: Math.min(seed.r1, r), c1: Math.min(seed.c1, c, seed.c1),
      r2: Math.max(seed.r2, r), c2: Math.max(seed.c2, c, seed.c2)
    })
    // Excel extends one axis at a time, whichever the mouse moved further on.
    const dRow = Math.abs(r - (r < seed.r1 ? seed.r1 : seed.r2))
    const dCol = Math.abs(c - (c < seed.c1 ? seed.c1 : seed.c2))
    const finalTarget = dRow >= dCol
      ? { r1: Math.min(seed.r1, r), c1: seed.c1, r2: Math.max(seed.r2, r), c2: seed.c2 }
      : { r1: seed.r1, c1: Math.min(seed.c1, c), r2: seed.r2, c2: Math.max(seed.c2, c) }
    ;(drag.current as { preview?: Range }).preview = finalTarget
    void target
    forceRedraw()
  }
  function onFillUp(e: MouseEvent): void {
    const d = drag.current as (null | { kind: 'fill'; seed: Range; preview?: Range })
    if (d && d.preview && (d.preview.r2 > d.seed.r2 || d.preview.r1 < d.seed.r1 || d.preview.c2 > d.seed.c2 || d.preview.c1 < d.seed.c1)) {
      cmd.fillHandleDrag(d.seed, d.preview, e.ctrlKey)
      setSel([d.preview], { r: sheet.selection.active.r, c: sheet.selection.active.c })
    }
    drag.current = null
    window.removeEventListener('mousemove', onFillDrag)
    window.removeEventListener('mouseup', onFillUp)
  }

  function onResizeMove(e: MouseEvent): void {
    const d = drag.current
    if (!d) return
    const rect = hostRef.current!.getBoundingClientRect()
    if (d.kind === 'resizeCol') {
      const dx = (e.clientX - rect.left - d.startX) / z
      const w = Math.max(2, Math.round(d.startW + dx))
      setResizeTip({ x: e.clientX - rect.left + 12, y: e.clientY - rect.top - 8, text: `Width: ${(w / 7).toFixed(2)} (${w} pixels)` })
    } else if (d.kind === 'resizeRow') {
      const dy = (e.clientY - rect.top - d.startY) / z
      const h = Math.max(2, Math.round(d.startH + dy))
      setResizeTip({ x: e.clientX - rect.left + 12, y: e.clientY - rect.top - 8, text: `Height: ${h}` })
    }
  }
  function onResizeUp(e: MouseEvent): void {
    const d = drag.current
    const rect = hostRef.current!.getBoundingClientRect()
    if (d && d.kind === 'resizeCol') {
      const dx = (e.clientX - rect.left - d.startX) / z
      const w = Math.max(2, Math.round(d.startW + dx))
      store.setColWidth(sheet, d.cols, w)
    } else if (d && d.kind === 'resizeRow') {
      const dy = (e.clientY - rect.top - d.startY) / z
      const h = Math.max(2, Math.round(d.startH + dy))
      store.setRowHeight(sheet, d.rows, h)
    }
    drag.current = null
    setResizeTip(null)
    window.removeEventListener('mousemove', onResizeMove)
    window.removeEventListener('mouseup', onResizeUp)
  }

  const pointedLen = useRef(0)
  const pointCur = useRef<{ ar: number; ac: number; r: number; c: number } | null>(null)
  function canPoint(): boolean {
    const e = getUi().edit
    if (!e || !e.text.startsWith('=')) return false
    if (pointCur.current) return true
    const before = e.text.slice(0, e.caret ?? e.text.length)
    return /[=(,+\-*/^&<>:;%{]\s*$/.test(before)
  }

  function insertRefIntoEdit(ar: number, ac: number, r: number, c: number): void {
    const e = getUi().edit
    if (!e) return
    const ref = ar === r && ac === c ? a1(r, c) : rangeA1({ r1: Math.min(ar, r), c1: Math.min(ac, c), r2: Math.max(ar, r), c2: Math.max(ac, c) })
    const base = e.text
    const caret = e.caret ?? base.length
    // Replace a trailing in-progress ref token (from a previous point-mode step) if present.
    const m = /(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)$/.exec(base.slice(0, caret))
    const start = m && pointedLen.current === m[1].length ? caret - m[1].length : caret
    pointedLen.current = ref.length
    const text = base.slice(0, start) + ref + base.slice(caret)
    setUi({ edit: { ...e, text, caret: start + ref.length } })
    store.setMode('Point')
  }

  function onDoubleClick(e: React.MouseEvent): void {
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    const rcol = findColBorder(x, y)
    if (rcol !== null && y < headerH) { autofitOneCol(rcol); return }
    const rrow = findRowBorder(x, y)
    if (rrow !== null && x < rowHeaderW) { store.setRowHeight(sheet, [rrow], DEFAULT_ROW_HEIGHT); return }
    if (inFillHandle(x, y)) {
      const seed = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
      let bottom = seed.r2
      const col = seed.c1 - 1 >= 0 ? seed.c1 - 1 : seed.c2 + 1
      while (bottom + 1 < MAX_ROWS && store.getRaw(sheet, bottom + 1, col) !== '') bottom++
      if (bottom > seed.r2) cmd.fillHandleDrag(seed, { ...seed, r2: bottom }, false)
      return
    }
    if (y < headerH || x < rowHeaderW) return
    tabAnchorCol.current = null
    const { r, c } = cellAt(x, y)
    beginEdit(r, c, 'Edit')
  }

  function onHostMouseMove(e: React.MouseEvent): void {
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    if (y < headerH || x < rowHeaderW) { if (hoverCell) setHoverCell(null); return }
    const { r, c } = cellAt(x, y)
    if (!hoverCell || hoverCell.r !== r || hoverCell.c !== c) setHoverCell({ r, c })
  }

  function autofitOneCol(c: number): void {
    const used2 = store.usedRange(sheet)
    let max = 40
    if (used2) for (let r = used2.r1; r <= Math.min(used2.r2, used2.r1 + 3000); r++) {
      const t = store.getDisplay(sheet, r, c).text
      if (t) max = Math.max(max, 7 * t.length + 12)
    }
    store.setColWidth(sheet, [c], max)
  }

  // ---------- keyboard ----------
  function onKeyDown(e: React.KeyboardEvent): void {
    if (e.target !== hostRef.current) return
    const editing = !!getUi().edit
    const mod = e.ctrlKey || e.metaKey
    if (!editing) {
      const extending = e.shiftKey || getUi().selectMode !== 'none'
      if (e.altKey && e.shiftKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) { outlineCommand(e.key === 'ArrowRight' ? 'group' : 'ungroup'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === '8') { setUi({ outlineHidden: !getUi().outlineHidden }); e.preventDefault(); return }
      // ---- Scroll Lock: arrows scroll the viewport, active cell doesn't move ----
      if (getUi().scrollLock && !mod && (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
        const el = scrollRef.current
        if (el) {
          if (e.key === 'ArrowDown') el.scrollTop += DEFAULT_ROW_HEIGHT * z
          else if (e.key === 'ArrowUp') el.scrollTop -= DEFAULT_ROW_HEIGHT * z
          else if (e.key === 'ArrowRight') el.scrollLeft += DEFAULT_COL_WIDTH * z
          else el.scrollLeft -= DEFAULT_COL_WIDTH * z
        }
        e.preventDefault(); return
      }
      if (e.key === 'ArrowDown') { tabAnchorCol.current = null; const j = mod || getUi().endMode; moveActive(1, 0, extending, { jump: j }); if (getUi().endMode) setUi({ endMode: false }); e.preventDefault(); return }
      if (e.key === 'ArrowUp') { tabAnchorCol.current = null; const j = mod || getUi().endMode; moveActive(-1, 0, extending, { jump: j }); if (getUi().endMode) setUi({ endMode: false }); e.preventDefault(); return }
      if (e.key === 'ArrowLeft') { tabAnchorCol.current = null; const j = mod || getUi().endMode; moveActive(0, -1, extending, { jump: j }); if (getUi().endMode) setUi({ endMode: false }); e.preventDefault(); return }
      if (e.key === 'ArrowRight') { tabAnchorCol.current = null; const j = mod || getUi().endMode; moveActive(0, 1, extending, { jump: j }); if (getUi().endMode) setUi({ endMode: false }); e.preventDefault(); return }
      // ---- Scroll Lock toggle ----
      if (e.key === 'ScrollLock') { setUi({ scrollLock: !getUi().scrollLock }); e.preventDefault(); return }
      // ---- End mode: arm, then the next arrow jumps once (see ArrowDown/Up/Left/Right above) ----
      if (e.key === 'End' && !mod) { setUi({ endMode: !getUi().endMode }); e.preventDefault(); return }
      // ---- Alt+Down: pick-from-list ----
      if (e.key === 'ArrowDown' && e.altKey && !mod) {
        // A cell with a List data validation opens its own dropdown instead, as in Excel.
        const a0 = sheet.selection.active
        const dv = getValidationForCell(sheet, a0.r, a0.c)
        if (dv && dv.allow === 'list' && dv.inCellDropdown !== false) {
          const rect = cellRect(a0.r, a0.c)
          const hostR = hostRef.current!.getBoundingClientRect()
          setListMenuPos({ x: hostR.left + rect.x, y: hostR.top + rect.y + rect.h })
          e.preventDefault(); return
        }
        const values = cmd.pickListValues()
        if (values.length) {
          const a = sheet.selection.active
          const rect = cellRect(a.r, a.c)
          const hostR = hostRef.current!.getBoundingClientRect()
          setUi({ pickList: { r: a.r, c: a.c, x: hostR.left + rect.x, y: hostR.top + rect.y + rect.h, values } })
        }
        e.preventDefault(); return
      }
      // ---- Ctrl+[ / Ctrl+] precedents/dependents, Ctrl+Shift+{ / } all levels ----
      if (mod && !e.shiftKey && e.key === '[') { cmd.selectPrecedents(false); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === ']') { cmd.selectDependents(false); e.preventDefault(); return }
      if (mod && e.shiftKey && e.key === '{') { cmd.selectPrecedents(true); e.preventDefault(); return }
      if (mod && e.shiftKey && e.key === '}') { cmd.selectDependents(true); e.preventDefault(); return }
      // ---- Ctrl+/: select the current array ----
      if (mod && !e.shiftKey && e.key === '/') { cmd.selectCurrentArray(); e.preventDefault(); return }
      // ---- Alt+;: select only visible cells ----
      if (e.altKey && e.key === ';') { cmd.selectVisibleCells(); e.preventDefault(); return }
      // ---- Ctrl+F2: Print in the backstage ----
      if (mod && e.key === 'F2') { setUi({ backstage: true, backstageTab: 'print' }); e.preventDefault(); return }
      // ---- Ctrl+Shift+PageUp/PageDown: group-select adjacent sheet tabs ----
      if (mod && e.shiftKey && (e.key === 'PageDown' || e.key === 'PageUp')) {
        const dir = e.key === 'PageDown' ? 1 : -1
        const cur = getUi().groupedSheets
        const base = cur && cur.includes(store.activeIndex) ? cur : [store.activeIndex]
        const next = Math.max(0, Math.min(store.sheets.length - 1, store.activeIndex + dir))
        const grouped = Array.from(new Set([...base, next])).sort((a, b) => a - b)
        setUi({ groupedSheets: grouped })
        store.setActiveSheet(next)
        e.preventDefault(); return
      }
      if (e.key === 'F8' && !e.shiftKey && !mod) { setUi({ selectMode: getUi().selectMode === 'extend' ? 'none' : 'extend' }); e.preventDefault(); return }
      if (e.key === 'F8' && e.shiftKey) { setUi({ selectMode: getUi().selectMode === 'add' ? 'none' : 'add' }); e.preventDefault(); return }
      if (e.key === 'Tab') {
        if (tabAnchorCol.current === null) tabAnchorCol.current = sheet.selection.active.c
        moveWithinOrGrid(0, e.shiftKey ? -1 : 1)
        e.preventDefault(); return
      }
      if (e.key === 'Enter') { commitEnterAfterTabAware(e.shiftKey); e.preventDefault(); return }
      if (e.key === 'Home') {
        tabAnchorCol.current = null
        const a = sheet.selection.active
        if (mod) { selectSingle(0, 0); ensureVisible(0, 0) } else { selectSingle(a.r, 0); ensureVisible(a.r, 0) }
        e.preventDefault(); return
      }
      if (e.key === 'End' && mod) {
        tabAnchorCol.current = null
        const u = store.usedRange(sheet) ?? { r1: 0, c1: 0, r2: 0, c2: 0 }
        selectSingle(u.r2, u.c2); ensureVisible(u.r2, u.c2); e.preventDefault(); return
      }
      if (e.key === 'PageDown' && !e.altKey && !mod) { pageMove(1); e.preventDefault(); return }
      if (e.key === 'PageUp' && !e.altKey && !mod) { pageMove(-1); e.preventDefault(); return }
      if (e.key === 'PageDown' && e.altKey) { moveActive(0, 10, e.shiftKey); e.preventDefault(); return }
      if (e.key === 'PageUp' && e.altKey) { moveActive(0, -10, e.shiftKey); e.preventDefault(); return }
      if (e.key === 'F2') { tabAnchorCol.current = null; const a = sheet.selection.active; beginEdit(a.r, a.c, 'Edit'); e.preventDefault(); return }
      if (e.key === 'Delete') { cmd.clearSelection('contents'); e.preventDefault(); return }
      if (e.key === 'Backspace') { const a = sheet.selection.active; cmd.clearSelection('contents'); beginEdit(a.r, a.c, 'Enter'); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'a') { cmd.selectAll(); e.preventDefault(); return }
      if (mod && e.key === ' ') { const cur = sheet.selection; const rg = cur.ranges[cur.ranges.length - 1]; setSel([{ r1: 0, c1: rg.c1, r2: MAX_ROWS - 1, c2: rg.c2 }], cur.active); e.preventDefault(); return }
      if (e.shiftKey && e.key === ' ') { const cur = sheet.selection; const rg = cur.ranges[cur.ranges.length - 1]; setSel([{ r1: rg.r1, c1: 0, r2: rg.r2, c2: MAX_COLS - 1 }], cur.active); e.preventDefault(); return }
      if (mod && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'e') { cmd.flashFillSelection(); e.preventDefault(); return }
      if (mod && e.shiftKey && e.key.toLowerCase() === 'l') { store.toggleAutoFilter(sheet); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'z') { store.undo(); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'y') { store.redo(); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'c') { cmd.copy(); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'x') { cmd.cut(); e.preventDefault(); return }
      if (mod && !e.altKey && e.key.toLowerCase() === 'v') { cmd.paste(); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'd') { cmd.fill('down'); e.preventDefault(); return }
      if (mod && e.key.toLowerCase() === 'r') { cmd.fill('right'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key.toLowerCase() === 'b') { toggleStyle('bold'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key.toLowerCase() === 'i') { toggleStyle('italic'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key.toLowerCase() === 'u') { toggleStyle('underline'); e.preventDefault(); return }
      if (mod && e.key === '1') { setUi({ dialog: { kind: 'formatCells' } }); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === '-') { cmd.deleteCellsDialog(); e.preventDefault(); return }
      if (mod && (e.key === '+' || (e.shiftKey && e.key === '='))) { cmd.insertCellsDialog(); e.preventDefault(); return }
      if (e.altKey && e.key === '=') { cmd.autoSum(); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === ';') { const a = sheet.selection.active; const d = new Date(); store.setRaw(sheet, a.r, a.c, `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`); e.preventDefault(); return }
      if (mod && e.shiftKey && (e.key === ':' || e.key === ';')) { const a = sheet.selection.active; const d = new Date(); store.setRaw(sheet, a.r, a.c, `${d.getHours() % 12 || 12}:${String(d.getMinutes()).padStart(2, '0')} ${d.getHours() >= 12 ? 'PM' : 'AM'}`); e.preventDefault(); return }
      if (e.key === 'Escape') { setUi({ selectMode: 'none' }); e.preventDefault(); return }
      // ---- number format shortcuts: Ctrl+Shift+~ ! @ # $ % ^ ----
      if (mod && e.key === '~') { cmd.applyNumFmtShortcut('general'); e.preventDefault(); return }
      if (mod && e.key === '!') { cmd.applyNumFmtShortcut('number'); e.preventDefault(); return }
      if (mod && e.key === '@') { cmd.applyNumFmtShortcut('time'); e.preventDefault(); return }
      if (mod && e.key === '#') { cmd.applyNumFmtShortcut('date'); e.preventDefault(); return }
      if (mod && e.key === '$') { cmd.applyNumFmtShortcut('currency'); e.preventDefault(); return }
      if (mod && e.key === '%') { cmd.applyNumFmtShortcut('percent'); e.preventDefault(); return }
      if (mod && e.key === '^') { cmd.applyNumFmtShortcut('scientific'); e.preventDefault(); return }
      // ---- quick borders: Ctrl+Shift+& / Ctrl+Shift+_ ----
      if (mod && e.key === '&') { cmd.applyOutlineBorder('box'); e.preventDefault(); return }
      if (mod && e.key === '_') { cmd.applyOutlineBorder('none'); e.preventDefault(); return }
      // ---- Ctrl+2/3/4/5: bold / italic / underline / strikethrough ----
      if (mod && !e.shiftKey && e.key === '2') { toggleStyle('bold'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === '3') { toggleStyle('italic'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === '4') { toggleStyle('underline'); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === '5') { const rg = sheet.selection.ranges; const st = store.getStyle(sheet, sheet.selection.active.r, sheet.selection.active.c); store.applyStyle(sheet, rg, { strike: !st?.strike }, 'Strikethrough'); e.preventDefault(); return }
      // ---- Ctrl+9/0 hide rows/columns, Ctrl+Shift+( / ) unhide ----
      if (mod && !e.shiftKey && e.key === '9') { cmd.hideRows(true); e.preventDefault(); return }
      if (mod && !e.shiftKey && e.key === '0') { cmd.hideColumns(true); e.preventDefault(); return }
      if (mod && e.key === '(') { cmd.hideRows(false); e.preventDefault(); return }
      if (mod && e.key === ')') { cmd.hideColumns(false); e.preventDefault(); return }
      // ---- Ctrl+` show formulas ----
      if (mod && e.key === '`') { setUi({ formulasVisible: !getUi().formulasVisible }); e.preventDefault(); return }
      // ---- sheet navigation: Ctrl+PgUp/PgDn switch sheet, Shift+F11 / Alt+Shift+F1 insert sheet ----
      if (mod && e.key === 'PageDown') { for (let k = store.activeIndex + 1; k < store.sheets.length; k++) if (!store.sheets[k].layout.sheetHidden) { store.setActiveSheet(k); break }; e.preventDefault(); return }
      if (mod && e.key === 'PageUp') { for (let k = store.activeIndex - 1; k >= 0; k--) if (!store.sheets[k].layout.sheetHidden) { store.setActiveSheet(k); break }; e.preventDefault(); return }
      if (e.key === 'F11' && e.shiftKey && !mod) { store.addSheet(); e.preventDefault(); return }
      if (e.key === 'F1' && e.altKey && e.shiftKey) { store.addSheet(); e.preventDefault(); return }
      // ---- Alt+F1: embedded chart of the default type (clustered column) from the selection ----
      if (e.key === 'F1' && e.altKey && !e.shiftKey && !mod) { insertChartFromSelection(store, COLUMN_BAR_GALLERY[0]); e.preventDefault(); return }
      // ---- F9 / Shift+F9 / Ctrl+Alt+F9: calculate now / active sheet / everything ----
      if (e.key === 'F9' && !(mod && !e.altKey)) { if (e.shiftKey) store.recalculateSheet(); else store.recalculate(); e.preventDefault(); return }
      // ---- Shift+F3: Insert Function ----
      if (e.key === 'F3' && e.shiftKey) { const a = sheet.selection.active; setUi({ edit: { r: a.r, c: a.c, text: '=', source: 'bar', mode: 'Enter' } }); e.preventDefault(); return }
      // ---- Ctrl+Alt+V: Paste Special ----
      if (mod && e.altKey && e.key.toLowerCase() === 'v') { setUi({ dialog: { kind: 'pasteSpecial' } }); e.preventDefault(); return }
      // ---- copy value/formula from the cell above ----
      if (mod && e.shiftKey && e.key === '"') { cmd.copyFromAbove('value'); e.preventDefault(); return }
      if (mod && e.key === "'") { cmd.copyFromAbove('formula'); e.preventDefault(); return }
      // ---- Ctrl+Shift+*: select current region ----
      if (mod && e.key === '*') { cmd.selectCurrentRegion(); e.preventDefault(); return }
      // ---- Shift+Backspace: collapse selection to the active cell ----
      if (e.shiftKey && e.key === 'Backspace') { cmd.collapseToActiveCell(); e.preventDefault(); return }
      // ---- Ctrl+Shift+Space: select all (approximates Excel's current-region-then-sheet escalation) ----
      if (mod && e.shiftKey && e.key === ' ') { cmd.selectAll(); e.preventDefault(); return }
      // ---- Ctrl+Shift+F / Ctrl+Shift+P: Format Cells > Font ----
      if (mod && e.shiftKey && (e.key.toLowerCase() === 'f' || e.key.toLowerCase() === 'p')) { setUi({ dialog: { kind: 'formatCells', props: { tab: 'Font' } } }); e.preventDefault(); return }
      // ---- Shift+F10 / ContextMenu key: open the right-click menu at the active cell ----
      if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') {
        const a = sheet.selection.active
        const rect = cellRect(a.r, a.c)
        const hostR = hostRef.current!.getBoundingClientRect()
        setCtxMenuKind('cell')
        openCtxMenu({ clientX: hostR.left + rect.x + rect.w / 2, clientY: hostR.top + rect.y + rect.h / 2 })
        e.preventDefault()
        return
      }
      if (e.key.length === 1 && !mod && !e.altKey) {
        const a = sheet.selection.active
        beginEdit(a.r, a.c, 'Enter', e.key)
        e.preventDefault()
        return
      }
      return
    }
  }

  function pageMove(dir: 1 | -1): void {
    const rowsPerPage = Math.max(1, Math.floor((size.h - headerH) / (DEFAULT_ROW_HEIGHT * z)))
    moveActive(dir * rowsPerPage, 0, false)
  }

  function toggleStyle(key: 'bold' | 'italic' | 'underline'): void {
    const rg = sheet.selection.ranges
    const first = store.getStyle(sheet, sheet.selection.active.r, sheet.selection.active.c)
    if (key === 'underline') store.applyStyle(sheet, rg, { underline: first?.underline ? undefined : 'single' })
    else store.applyStyle(sheet, rg, { [key]: !first?.[key] } as Partial<CellStyle>)
  }

  /** Enter: normally moves down one row in place; but after a Tab-run in a single-cell
   *  selection, returns to the column where that Tab-run started (Excel's Enter-after-Tab). */
  function commitEnterAfterTabAware(shift: boolean): void {
    const anchor = tabAnchorCol.current
    if (anchor !== null && !shift) {
      tabAnchorCol.current = null
      const a = sheet.selection.active
      const r = Math.min(MAX_ROWS - 1, a.r + 1)
      selectSingle(r, anchor)
      ensureVisible(r, anchor)
    } else {
      moveWithinOrGrid(shift ? -1 : 1, 0)
    }
  }

  function moveWithinOrGrid(dr: number, dc: number): void {
    const cur = sheet.selection
    const rg = normRange(cur.ranges[cur.ranges.length - 1])
    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2
    if (!multi) { moveActive(dr, dc, false); return }
    let { r, c } = cur.active
    if (dc !== 0) {
      c += dc
      if (c > rg.c2) { c = rg.c1; r++; if (r > rg.r2) r = rg.r1 }
      else if (c < rg.c1) { c = rg.c2; r--; if (r < rg.r1) r = rg.r2 }
    } else {
      r += dr
      if (r > rg.r2) { r = rg.r1; c++; if (c > rg.c2) c = rg.c1 }
      else if (r < rg.r1) { r = rg.r2; c--; if (c < rg.c1) c = rg.c2 }
    }
    setSel(cur.ranges, { r, c })
    ensureVisible(r, c)
  }

  // ---------- editor textarea handlers ----------
  function onEditorKeyDown(e: React.KeyboardEvent<HTMLTextAreaElement>): void {
    const edit = getUi().edit!
    if (e.key === 'Escape') { tabAnchorCol.current = null; cancelEdit(); e.preventDefault(); hostRef.current?.focus(); return }
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && !e.altKey) {
      commitEditToSelection()
      hostRef.current?.focus()
      e.preventDefault()
      return
    }
    if (e.key === 'Enter' && !e.altKey) {
      commitEdit()
      commitEnterAfterTabAware(e.shiftKey)
      hostRef.current?.focus()
      e.preventDefault()
      return
    }
    if (e.key === 'Tab') {
      if (tabAnchorCol.current === null) tabAnchorCol.current = getUi().edit!.c
      commitEdit()
      moveWithinOrGrid(0, e.shiftKey ? -1 : 1)
      hostRef.current?.focus()
      e.preventDefault()
      return
    }
    const isArrow = e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight'
    if (!isArrow && e.key !== 'Shift' && e.key !== 'Control' && e.key !== 'F4') { pointCur.current = null; pointedLen.current = 0 }
    // Keyboard Point mode: arrows move a reference cursor and write its address into the formula.
    if (isArrow && edit.mode === 'Enter' && canPoint()) {
      const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0
      const dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0
      const p0 = pointCur.current ?? { ar: edit.r, ac: edit.c, r: edit.r, c: edit.c }
      const r = Math.max(0, Math.min(MAX_ROWS - 1, p0.r + dr))
      const c = Math.max(0, Math.min(MAX_COLS - 1, p0.c + dc))
      const next = e.shiftKey && pointCur.current ? { ...p0, r, c } : { ar: r, ac: c, r, c }
      pointCur.current = next
      insertRefIntoEdit(next.ar, next.ac, next.r, next.c)
      ensureVisible(r, c)
      e.preventDefault()
      return
    }
    if (edit.mode === 'Enter' && isArrow) {
      tabAnchorCol.current = null
      commitEdit()
      const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0
      const dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0
      moveActive(dr, dc, false)
      hostRef.current?.focus()
      e.preventDefault()
      return
    }
    // F4: while editing a formula, cycles the $-lock style of the reference at/before the caret
    // (A1 -> $A$1 -> A$1 -> $A1 -> A1), Excel's "toggle absolute/relative reference".
    if (e.key === 'F4' && !e.ctrlKey && !e.altKey) {
      const ta = e.target as HTMLTextAreaElement
      const text = edit.text
      const caret = ta.selectionEnd ?? edit.caret ?? text.length
      const re = /\$?[A-Za-z]{1,3}\$?\d{1,7}/g
      let m: RegExpExecArray | null
      let hit: RegExpExecArray | null = null
      while ((m = re.exec(text))) { if (m.index < caret && m.index + m[0].length >= caret) { hit = m; break } }
      if (hit) {
        const ref = hit[0]
        const colAbs = ref.startsWith('$')
        const rest = colAbs ? ref.slice(1) : ref
        const dollarBeforeRow = rest.indexOf('$')
        const rowAbs = dollarBeforeRow >= 0
        const col = (rowAbs ? rest.slice(0, dollarBeforeRow) : rest.replace(/\d+$/, ''))
        const row = rowAbs ? rest.slice(dollarBeforeRow + 1) : rest.replace(/^[A-Za-z]+/, '')
        const next =
          !colAbs && !rowAbs ? `$${col}$${row}` :
          colAbs && rowAbs ? `${col}$${row}` :
          !colAbs && rowAbs ? `$${col}${row}` :
          `${col}${row}`
        const newText = text.slice(0, hit.index) + next + text.slice(hit.index + ref.length)
        const newCaret = hit.index + next.length
        setUi({ edit: { ...edit, text: newText, caret: newCaret } })
        requestAnimationFrame(() => { ta.setSelectionRange(newCaret, newCaret) })
        e.preventDefault()
      }
      return
    }
  }
  function onEditorChange(e: React.ChangeEvent<HTMLTextAreaElement>): void {
    const edit = getUi().edit!
    const native = e.nativeEvent as InputEvent
    const val = e.target.value
    // A Tab or newline character can arrive as a plain text insertion (inputType "insertText")
    // instead of a keydown — e.g. IME composition, accessibility tools, or automation that
    // synthesizes text input directly. A real Tab/Enter keypress never reaches here as inserted
    // text (its keydown handler commits and moves first), and a genuine paste is reported as
    // "insertFromPaste", so this only fires for that synthetic-insert case — where we replay the
    // same commit-and-move behavior as an actual Tab/Enter keystroke instead of embedding the
    // control character in the cell's text.
    if (native?.inputType === 'insertText' && (native.data === '\t' || native.data === '\n') && val.endsWith(native.data)) {
      const text = val.slice(0, -1)
      setUi({ edit: { ...edit, text, caret: text.length } })
      if (native.data === '\t') {
        if (tabAnchorCol.current === null) tabAnchorCol.current = edit.c
        commitEdit()
        moveWithinOrGrid(0, 1)
      } else {
        commitEdit()
        commitEnterAfterTabAware(false)
      }
      hostRef.current?.focus()
      return
    }
    setUi({ edit: { ...edit, text: val, caret: e.target.selectionStart ?? val.length } })
    store.setMode(edit.mode)
  }

  // ---------- context menu ----------
  function onContextMenu(e: React.MouseEvent): void {
    e.preventDefault()
    const rect = hostRef.current!.getBoundingClientRect()
    const x = e.clientX - rect.left, y = e.clientY - rect.top
    if (showHead && y < headerH && x >= rowHeaderW) {
      const c = colAt(x)
      if (!selectedColsContains(c)) { const rg: Range = { r1: 0, c1: c, r2: MAX_ROWS - 1, c2: c }; setSel([rg], { r: 0, c }) }
      setCtxMenuKind('col')
    } else if (showHead && x < rowHeaderW && y >= headerH) {
      const r = rowAt(y)
      if (!selectedRowsContains(r)) { const rg: Range = { r1: r, c1: 0, r2: r, c2: MAX_COLS - 1 }; setSel([rg], { r, c: 0 }) }
      setCtxMenuKind('row')
    } else if (x >= rowHeaderW && y >= headerH) {
      const { r, c } = cellAt(x, y)
      if (!sheet.selection.ranges.some((rg) => inRange(rg, r, c))) selectSingle(r, c)
      setCtxMenuKind('cell')
    } else return
    openCtxMenu(e)
  }

  const PASTE_OPTS: { k: cmd.PasteKind; label: string; icon: React.ReactNode; badge?: string }[] = [
    { k: 'all', label: 'Paste', icon: <ClipboardPaste20Regular /> },
    { k: 'values', label: 'Values', icon: <ClipboardTaskRegular />, badge: '123' },
    { k: 'formulas', label: 'Formulas', icon: <ClipboardCodeRegular />, badge: 'fx' },
    { k: 'transpose', label: 'Transpose', icon: <ArrowSwap20Regular /> },
    { k: 'formats', label: 'Formatting', icon: <PaintBrush20Regular /> },
    { k: 'link', label: 'Paste Link', icon: <ClipboardLink20Regular /> }
  ]
  const pasteOptionsStrip = (close: () => void) => (
    <div className="xl-paste-strip">
      {PASTE_OPTS.map(({ k, label, icon, badge }) => (
        <button key={k} className="xl-paste-icon" title={label} onClick={() => { close(); cmd.paste(k) }}>
          {icon}
          {badge && <span className="xl-paste-badge">{badge}</span>}
        </button>
      ))}
    </div>
  )

  const miniToolbar = () => {
    const style = store.getStyle(sheet, sheet.selection.active.r, sheet.selection.active.c) ?? {}
    const setStyle = (patch: Partial<CellStyle>) => store.applyStyle(sheet, sheet.selection.ranges, patch, 'Format Cells')
    return (
      <div className="xl-mini-toolbar">
        <span className="xl-mini-size">{style.fontSize ?? 11}</span>
        <MiniBtn icon={<FontIncrease20Regular />} title="Increase Font Size" onClick={() => setStyle({ fontSize: (style.fontSize ?? 11) + 1 })} />
        <MiniBtn icon={<FontDecrease20Regular />} title="Decrease Font Size" onClick={() => setStyle({ fontSize: Math.max(1, (style.fontSize ?? 11) - 1) })} />
        <span className="xl-mini-sep" />
        <MiniBtn icon={<TextBold20Regular />} title="Bold" active={!!style.bold} onClick={() => toggleStyle('bold')} />
        <MiniBtn icon={<TextItalic20Regular />} title="Italic" active={!!style.italic} onClick={() => toggleStyle('italic')} />
        <MiniBtn icon={<TextUnderline20Regular />} title="Underline" active={!!style.underline} onClick={() => toggleStyle('underline')} />
        <MiniBtn icon={<BorderAll20Regular />} title="All Borders" onClick={() => cmd.applyBorderPreset('all')} />
        <MiniBtn icon={<PaintBucket20Regular />} title="Fill Color" onClick={() => setStyle({ fill: '#FFFF00' })} />
        <MiniBtn icon={<TextFont20Regular />} title="Font Color" onClick={() => setStyle({ color: '#FF0000' })} />
        <span className="xl-mini-sep" />
        <MiniBtn icon={<AlignLeft20Regular />} title="Align Left" active={style.hAlign === 'left'} onClick={() => setStyle({ hAlign: 'left' })} />
        <MiniBtn icon={<AlignCenterHorizontal20Regular />} title="Center" active={style.hAlign === 'center'} onClick={() => setStyle({ hAlign: 'center' })} />
        <MiniBtn icon={<AlignRight20Regular />} title="Align Right" active={style.hAlign === 'right'} onClick={() => setStyle({ hAlign: 'right' })} />
        <span className="xl-mini-sep" />
        <MiniBtn icon={<TextPercent20Regular />} title="Percent Style" onClick={() => setStyle({ numFmt: '0%' })} />
        <MiniBtn icon={<TextIndentDecrease20Regular />} title="Decrease Decimal" onClick={() => cmd.adjustDecimal(-1)} />
        <MiniBtn icon={<TextIndentIncrease20Regular />} title="Increase Decimal" onClick={() => cmd.adjustDecimal(1)} />
      </div>
    )
  }

  const activeKey = sheet.selection.active.r + ',' + sheet.selection.active.c
  const hasNote = !!sheet.layout.notes[activeKey]
  const hasComment = !!sheet.layout.comments[activeKey]
  const pv = (() => {
    const a = sheet.selection.active
    return (sheet.layout.pivots ?? []).find((p) => p.lastSize && a.r >= p.anchor.r && a.r < p.anchor.r + p.lastSize.rows && a.c >= p.anchor.c && a.c < p.anchor.c + p.lastSize.cols)
  })()
  const cellMenu: MenuItem[] = [
    ...(pv
      ? [
          { label: '&Refresh', onClick: () => { const e = renderPivot(store, sheet, pv); if (e) setUi({ dialog: { kind: 'message', props: { text: e } } }) } } as MenuItem,
          { label: 'Show Field &List', onClick: () => setUi({ dialog: { kind: 'pivot', props: { edit: { sheetUid: sheet.uid, id: pv.id } } } }) } as MenuItem,
          'separator' as const
        ]
      : []),
    { label: 'Cu&t', icon: <Cut20Regular />, shortcut: 'Ctrl+X', onClick: () => cmd.cut() },
    { label: '&Copy', icon: <Copy20Regular />, shortcut: 'Ctrl+C', onClick: () => cmd.copy() },
    { label: 'Paste Options:', render: pasteOptionsStrip },
    { label: 'Paste &Special...', icon: <ClipboardPaste20Regular />, onClick: () => setUi({ dialog: { kind: 'pasteSpecial' } }) },
    'separator',
    { label: 'Smart &Lookup', icon: <SquareHintSparkles20Regular /> },
    'separator',
    { label: '&Insert...', icon: <TableAdd20Regular />, onClick: () => cmd.insertCellsDialog() },
    { label: '&Delete...', icon: <Delete20Regular />, onClick: () => cmd.deleteCellsDialog() },
    { label: 'Clear Co&ntents', icon: <Eraser20Regular />, shortcut: 'Del', onClick: () => cmd.clearSelection('contents') },
    { label: '&Translate', icon: <Translate20Regular /> },
    { label: '&Quick Analysis', icon: <SquareHintSparkles20Regular /> },
    'separator',
    {
      label: 'Filt&er', icon: <Filter20Regular />, submenu: [
        {
          label: "Filter by Selected Cell's Value",
          onClick: () => {
            const { r, c } = sheet.selection.active
            if (!sheet.layout.autoFilter) store.toggleAutoFilter(sheet)
            if (sheet.layout.autoFilter) store.setColumnFilter(sheet, c, { selected: [store.getDisplay(sheet, r, c).text] })
          }
        },
        { label: 'Clear Filter', onClick: () => store.clearAllFilters(sheet), disabled: !sheet.layout.autoFilter }
      ]
    },
    { label: 'S&ort', icon: <ArrowSort20Regular />, submenu: [{ label: 'Smallest to Largest', onClick: () => cmd.sortSelection('asc') }, { label: 'Largest to Smallest', onClick: () => cmd.sortSelection('desc') }] },
    'separator',
    { label: 'Get Data from Table/Range...' },
    hasComment
      ? { label: '&Edit Comment', icon: <Comment20Regular />, onClick: () => setUi({ commentPanel: { r: sheet.selection.active.r, c: sheet.selection.active.c } }) }
      : { label: 'New Co&mment', icon: <Comment20Regular />, onClick: () => setUi({ commentPanel: { r: sheet.selection.active.r, c: sheet.selection.active.c } }) },
    hasNote
      ? { label: '&Edit Note', icon: <Note20Regular />, onClick: () => setUi({ noteEditor: { r: sheet.selection.active.r, c: sheet.selection.active.c } }) }
      : { label: 'New N&ote', icon: <Note20Regular />, onClick: () => setUi({ noteEditor: { r: sheet.selection.active.r, c: sheet.selection.active.c } }) },
    ...(hasNote ? [{ label: '&Delete Note', icon: <Delete20Regular />, onClick: () => store.deleteNote(sheet, sheet.selection.active.r, sheet.selection.active.c) } as MenuItem] : []),
    'separator',
    { label: '&Format Cells...', icon: <TableCellEdit20Regular />, onClick: () => setUi({ dialog: { kind: 'formatCells' } }) },
    {
      label: 'Pic&k From Drop-down List...',
      onClick: () => {
        const values = cmd.pickListValues()
        const a = sheet.selection.active
        const rect = cellRect(a.r, a.c)
        const hostR = hostRef.current!.getBoundingClientRect()
        if (values.length) setUi({ pickList: { r: a.r, c: a.c, x: hostR.left + rect.x, y: hostR.top + rect.y + rect.h, values } })
      }
    },
    { label: 'Define N&ame...', onClick: () => setUi({ dialog: { kind: 'defineName' } }) },
    ...(linkAt(sheet, sheet.selection.active.r, sheet.selection.active.c)
      ? [
          { label: '&Edit Hyperlink...', icon: <Link20Regular />, onClick: () => setUi({ dialog: { kind: 'hyperlink' } }) } as MenuItem,
          { label: '&Open Hyperlink', onClick: () => { const l = linkAt(sheet, sheet.selection.active.r, sheet.selection.active.c); if (l) void followLink(l) } } as MenuItem,
          { label: '&Remove Hyperlink', onClick: () => cmd.clearSelection('hyperlinks') } as MenuItem
        ]
      : [{ label: 'L&ink', icon: <Link20Regular />, shortcut: 'Ctrl+K', onClick: () => setUi({ dialog: { kind: 'hyperlink' } }) } as MenuItem]),
    { label: 'Show Changes' }
  ]
  const colMenu: MenuItem[] = [
    { label: 'Cu&t', icon: <Cut20Regular />, shortcut: 'Ctrl+X', onClick: () => cmd.cut() },
    { label: '&Copy', icon: <Copy20Regular />, shortcut: 'Ctrl+C', onClick: () => cmd.copy() },
    { label: 'Paste Options:', render: pasteOptionsStrip },
    { label: 'Paste &Special...', icon: <ClipboardPaste20Regular />, onClick: () => setUi({ dialog: { kind: 'pasteSpecial' } }) },
    'separator',
    { label: '&Insert', icon: <TableAdd20Regular />, onClick: () => cmd.insertSheetColumns() },
    { label: '&Delete', icon: <Delete20Regular />, onClick: () => cmd.deleteSheetColumns() },
    { label: 'Clear Co&ntents', icon: <Eraser20Regular />, onClick: () => cmd.clearSelection('contents') },
    'separator',
    { label: '&Format Cells...', icon: <TableCellEdit20Regular />, onClick: () => setUi({ dialog: { kind: 'formatCells' } }) },
    { label: '&Column Width...', icon: <ColumnTriple20Regular />, onClick: () => setUi({ dialog: { kind: 'columnWidth' } }) },
    { label: '&Hide', onClick: () => cmd.hideColumns(true) },
    { label: '&Unhide', onClick: () => cmd.hideColumns(false) }
  ]
  const rowMenu: MenuItem[] = [
    { label: 'Cu&t', icon: <Cut20Regular />, shortcut: 'Ctrl+X', onClick: () => cmd.cut() },
    { label: '&Copy', icon: <Copy20Regular />, shortcut: 'Ctrl+C', onClick: () => cmd.copy() },
    { label: 'Paste Options:', render: pasteOptionsStrip },
    { label: 'Paste &Special...', icon: <ClipboardPaste20Regular />, onClick: () => setUi({ dialog: { kind: 'pasteSpecial' } }) },
    'separator',
    { label: '&Insert', icon: <TableAdd20Regular />, onClick: () => cmd.insertSheetRows() },
    { label: '&Delete', icon: <TableDeleteRow20Regular />, onClick: () => cmd.deleteSheetRows() },
    { label: 'Clear Co&ntents', icon: <Eraser20Regular />, onClick: () => cmd.clearSelection('contents') },
    'separator',
    { label: '&Format Cells...', icon: <TableCellEdit20Regular />, onClick: () => setUi({ dialog: { kind: 'formatCells' } }) },
    { label: '&Row Height...', icon: <ColumnTriple20Regular />, onClick: () => setUi({ dialog: { kind: 'rowHeight' } }) },
    { label: '&Hide', onClick: () => cmd.hideRows(true) },
    { label: '&Unhide', onClick: () => cmd.hideRows(false) }
  ]

  const edit = ui.edit

  const activeCell = sheet.selection.active
  const dvRule = getValidationForCell(sheet, activeCell.r, activeCell.c)
  const isListValidation = dvRule?.allow === 'list' && dvRule.inCellDropdown !== false
  const listItems = isListValidation && dvRule ? getListItems(store, sheet, dvRule) : []
  const activeRc = isListValidation ? cellRect(activeCell.r, activeCell.c) : null
  const showDropdownBtn =
    isListValidation &&
    !ui.edit &&
    activeRc !== null &&
    activeRc.x + activeRc.w > rowHeaderW &&
    activeRc.y + activeRc.h > headerH &&
    activeRc.x < size.w &&
    activeRc.y < size.h

  return (
    <div
      className="grid-host"
      ref={hostRef}
      tabIndex={0}
      onMouseDown={onMouseDown}
      onMouseMove={onHostMouseMove}
      onMouseLeave={() => setHoverCell(null)}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
      onKeyDown={onKeyDown}
      style={{ outline: 'none', cursor: 'cell' }}
    >
      <div ref={scrollRef} className="xl-scroll" onScroll={onScroll}>
        <div style={{ width: sizerW, height: sizerH }} />
      </div>
      <canvas ref={canvasRef} className="xl-canvas" style={{ width: size.w, height: size.h }} />
      <ChartOverlay
        sheet={sheet}
        geo={{ colsAxis, rowsAxis, frozen, scroll, z, rowHeaderW, headerH, viewport: size }}
      />
      {edit && edit.source === 'cell' && (
        <textarea
          ref={editorRef}
          className="xl-cell-editor"
          value={edit.text}
          autoFocus
          wrap="off"
          onChange={onEditorChange}
          onKeyDown={onEditorKeyDown}
          onBlur={() => { if (getUi().edit) commitEdit() }}
        />
      )}
      {showDropdownBtn && activeRc && (
        <button
          type="button"
          className="xl-dv-dropdown-btn"
          style={{
            position: 'absolute',
            left: Math.round(activeRc.x + activeRc.w - 18),
            top: Math.round(activeRc.y),
            width: 18,
            height: Math.round(activeRc.h),
            zIndex: 4
          }}
          title="List options"
          onMouseDown={(e) => {
            e.stopPropagation()
            const rect = hostRef.current?.getBoundingClientRect()
            const clientX = (rect?.left || 0) + activeRc.x
            const clientY = (rect?.top || 0) + activeRc.y + activeRc.h
            setListMenuPos(listMenuPos ? null : { x: clientX, y: clientY })
          }}
        >
          ▾
        </button>
      )}
      {listMenuPos && (
        <Menu
          x={listMenuPos.x}
          y={listMenuPos.y}
          onClose={() => setListMenuPos(null)}
          items={listItems.map((item) => ({
            label: item,
            onClick: () => {
              store.setRaw(sheet, activeCell.r, activeCell.c, item)
              setListMenuPos(null)
            }
          }))}
        />
      )}
      {resizeTip && <div className="xl-resize-tip" style={{ left: resizeTip.x, top: resizeTip.y }}>{resizeTip.text}</div>}
      {ctxMenu && (
        <Menu
          x={ctxMenu.x}
          y={ctxMenu.y}
          onClose={closeCtxMenu}
          items={ctxMenuKind === 'cell' ? cellMenu : ctxMenuKind === 'col' ? colMenu : rowMenu}
          header={ctxMenuKind === 'cell' ? miniToolbar() : undefined}
        />
      )}
      {filterMenu && sheet.layout.autoFilter && (
        <AutoFilterMenu sheet={sheet} col={filterMenu.col} x={filterMenu.x} y={filterMenu.y} onClose={() => setFilterMenu(null)} />
      )}
      {ui.noteEditor && (
        <NoteBox sheet={sheet} r={ui.noteEditor.r} c={ui.noteEditor.c} rect={cellRect(ui.noteEditor.r, ui.noteEditor.c)} onClose={closeNoteEditor} />
      )}
      {ui.commentPanel && (
        <CommentBox sheet={sheet} r={ui.commentPanel.r} c={ui.commentPanel.c} rect={cellRect(ui.commentPanel.r, ui.commentPanel.c)} onClose={closeCommentPanel} />
      )}
      {ui.pickList && (
        <Menu
          x={ui.pickList.x}
          y={ui.pickList.y}
          onClose={() => setUi({ pickList: null })}
          items={ui.pickList.values.map((v) => ({
            label: v || '(blank)',
            onClick: () => { store.setRaw(sheet, ui.pickList!.r, ui.pickList!.c, v); setUi({ pickList: null }); hostRef.current?.focus() }
          }))}
        />
      )}
      {!ui.noteEditor && Object.entries(sheet.layout.notes).filter(([k, n]) => n.visible && !(hoverCell && k === hoverCell.r + ',' + hoverCell.c)).slice(0, 200).map(([k]) => {
        const [nr, nc] = k.split(',').map(Number)
        const rc = cellRect(nr, nc)
        if (rc.x + rc.w < 0 || rc.y + rc.h < 0 || rc.x > size.w || rc.y > size.h) return null
        return <NoteTooltip key={'n' + k} sheet={sheet} r={nr} c={nc} rect={rc} />
      })}
      {hoverCell && !ui.edit && !ui.noteEditor && !filterMenu &&
        sheet.layout.notes[hoverCell.r + ',' + hoverCell.c] && (
          <NoteTooltip sheet={sheet} r={hoverCell.r} c={hoverCell.c} rect={cellRect(hoverCell.r, hoverCell.c)} />
        )}
      <GridDialogs sheet={sheet} />
    </div>
  )
}

// ---------- grid-owned dialogs ----------
function GridDialogs({ sheet }: { sheet: SheetState }) {
  const ui = useUi()
  const dlg = ui.dialog
  const [widthVal, setWidthVal] = useState('8.43')
  const [heightVal, setHeightVal] = useState('15')
  useEffect(() => {
    if (dlg?.kind === 'columnWidth') setWidthVal(String(+(store.colWidth(sheet, sheet.selection.active.c) / 7).toFixed(2)))
    if (dlg?.kind === 'rowHeight') setHeightVal(String(store.rowHeight(sheet, sheet.selection.active.r)))
  }, [dlg, sheet])
  if (!dlg) return null
  const close = () => setUi({ dialog: null })

  if (dlg.kind === 'insertCells') {
    return (
      <Dialog title="Insert" onCancel={close} onOk={() => { const v = (window as unknown as { __insertKind?: 'right' | 'down' | 'row' | 'col' }).__insertKind ?? 'right'; cmd.insertShift(v); close() }}>
        <InsertDeleteRadios kind="insert" />
      </Dialog>
    )
  }
  if (dlg.kind === 'deleteCells') {
    return (
      <Dialog title="Delete" onCancel={close} onOk={() => { const v = (window as unknown as { __deleteKind?: 'left' | 'up' | 'row' | 'col' }).__deleteKind ?? 'up'; cmd.deleteShift(v); close() }}>
        <InsertDeleteRadios kind="delete" />
      </Dialog>
    )
  }
  if (dlg.kind === 'columnWidth') {
    return (
      <Dialog title="Column Width" onCancel={close} onOk={() => { store.setColWidth(sheet, [...selCols(sheet)], Math.round(parseFloat(widthVal) * 7)); close() }}>
        <div className="xl-field-row"><label>Column width:</label><input autoFocus value={widthVal} onChange={(e) => setWidthVal(e.target.value)} /></div>
      </Dialog>
    )
  }
  if (dlg.kind === 'rowHeight') {
    return (
      <Dialog title="Row Height" onCancel={close} onOk={() => { store.setRowHeight(sheet, [...selRows(sheet)], parseFloat(heightVal)); close() }}>
        <div className="xl-field-row"><label>Row height:</label><input autoFocus value={heightVal} onChange={(e) => setHeightVal(e.target.value)} /></div>
      </Dialog>
    )
  }
  if (dlg.kind === 'pasteSpecial') {
    return (
      <Dialog title="Paste Special" onCancel={close} onOk={() => { cmd.paste((window as unknown as { __pasteKind?: cmd.PasteKind }).__pasteKind ?? 'all'); close() }} width={420}>
        <PasteSpecialBody />
      </Dialog>
    )
  }
  return null
}

function selCols(sheet: SheetState): Set<number> {
  const s = new Set<number>()
  sheet.selection.ranges.forEach((rg) => { const n = normRange(rg); for (let c = n.c1; c <= n.c2; c++) s.add(c) })
  return s
}
function selRows(sheet: SheetState): Set<number> {
  const s = new Set<number>()
  sheet.selection.ranges.forEach((rg) => { const n = normRange(rg); for (let r = n.r1; r <= n.r2; r++) s.add(r) })
  return s
}

function InsertDeleteRadios({ kind }: { kind: 'insert' | 'delete' }) {
  const opts = kind === 'insert'
    ? [['right', 'Shift cells right'], ['down', 'Shift cells down'], ['row', 'Entire row'], ['col', 'Entire column']]
    : [['left', 'Shift cells left'], ['up', 'Shift cells up'], ['row', 'Entire row'], ['col', 'Entire column']]
  const [val, setVal] = useState(opts[0][0])
  useEffect(() => { (window as unknown as Record<string, unknown>)[kind === 'insert' ? '__insertKind' : '__deleteKind'] = val }, [val, kind])
  return (
    <div className="xl-radio-col">
      {opts.map(([v, label]) => (
        <label key={v}><input type="radio" name={kind} checked={val === v} onChange={() => setVal(v)} /> {label}</label>
      ))}
    </div>
  )
}

function PasteSpecialBody() {
  const [val, setVal] = useState<cmd.PasteKind>('all')
  useEffect(() => { (window as unknown as Record<string, unknown>).__pasteKind = val }, [val])
  const pasteOpts: [cmd.PasteKind, string][] = [
    ['all', 'All'], ['formulas', 'Formulas'], ['values', 'Values'], ['formats', 'Formats'],
    ['columnWidths', 'Column widths'], ['noBorders', 'All except borders'], ['link', 'Formulas and number formats']
  ]
  return (
    <div className="xl-paste-special">
      <div className="xl-radio-col">
        <b>Paste</b>
        {pasteOpts.map(([v, label]) => (
          <label key={v}><input type="radio" name="pastekind" checked={val === v} onChange={() => setVal(v)} /> {label}</label>
        ))}
      </div>
      <div className="xl-radio-col">
        <b>Operation</b>
        <label><input type="radio" name="op" defaultChecked /> None</label>
        <label><input type="checkbox" onChange={(e) => setVal(e.target.checked ? 'transpose' : 'all')} /> Transpose</label>
      </div>
    </div>
  )
}

/** Shortest-first Excel General fallbacks for a number too wide for its column. */
export function fitGeneral(v: number, fits: (t: string) => boolean): string | null {
  const abs = Math.abs(v)
  if (!Number.isInteger(v) && abs < 1e11) {
    const intDigits = Math.max(1, Math.floor(Math.log10(Math.max(abs, 1))) + 1)
    for (let d = 9 - intDigits; d >= 0; d--) {
      const t = String(Number(v.toFixed(d)))
      if (fits(t)) return t
    }
  }
  if (abs >= 1e-9) {
    for (let d = 5; d >= 0; d--) {
      const [m, e] = v.toExponential(d).split('e')
      const exp = Number(e)
      const t = `${Number(m)}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`
      if (fits(t)) return t
    }
  }
  return null
}

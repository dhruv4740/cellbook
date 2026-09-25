// The workbook store: single source of truth for the app.
// Cell contents and formula evaluation live in HyperFormula; styles, layout and selection live here.
// Every mutation goes through `commit`, which records an undo entry and notifies subscribers.
import { HyperFormula, DetailedCellError } from 'hyperformula'
import { useSyncExternalStore } from 'react'
import {
  CellStyle, CellValue, Range, Selection, SheetLayout, cellKey, normRange,
  DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, MAX_ROWS, MAX_COLS, FilterCriteria, NoteData, CommentThread
} from './types'
import { formatValue, detectInputFormat, FormattedValue } from '../format/numfmt'
import { computeFilteredRows } from '../engine/filter'
import { shiftFormula } from '../engine/shiftFormula'
import { DEFAULT_AUTHOR, defaultLayout } from './defaults'
import type { StoreOp, SyncMessage } from './sync'
import { applySyncMessage } from './sync'

// Re-exported so existing `import { DEFAULT_AUTHOR, defaultLayout } from '../model/store'` call
// sites keep working; the actual definitions live in defaults.ts (no HyperFormula/React), which
// the fast xlsx reader/worker import directly instead of pulling in this whole heavy module.
export { DEFAULT_AUTHOR, defaultLayout }

export interface SheetState {
  uid: number // stable id, independent of HyperFormula's sheet id
  name: string
  styles: Map<string, CellStyle>
  layout: SheetLayout
  selection: Selection
}

/** Plain, serializable workbook used by file import/export. */
export interface WorkbookData {
  sheets: {
    name: string
    cells: Record<string, { raw?: string; style?: CellStyle }> // key 'r,c'
    layout?: Partial<SheetLayout>
  }[]
  activeSheet?: number
}

interface UndoEntry {
  label: string
  undo: () => void
  redo: () => void
}

function defaultSelection(): Selection {
  return { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }
}

let nextUid = 1

// US-locale input parsing, matching Excel en-US.
import { registerExtraFunctions, BOOLEAN_NAMES } from '../engine/extraFunctions'
registerExtraFunctions()

const HF_CONFIG = {
  maxRows: 1048576,
  maxColumns: 16384,
  licenseKey: 'gpl-v3',
  undoLimit: 0,
  useArrayArithmetic: true,
  dateFormats: ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY-MM-DD', 'YYYY/MM/DD'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  currencySymbol: ['$'],
  nullDate: { year: 1899, month: 12, day: 30 }
}

export class WorkbookStore {
  hf: HyperFormula
  sheets: SheetState[] = []
  activeIndex = 0
  fileName = 'Book1'
  filePath: string | null = null
  dirty = false
  version = 0
  /** Cells shown with marching ants after Copy/Cut; null when nothing is on the clipboard. */
  clipboard: { sheetUid: number; range: Range; cut: boolean; text?: string } | null = null
  /** Transient UI state any component may read (e.g. edit mode for the status bar). */
  mode: 'Ready' | 'Enter' | 'Edit' | 'Point' = 'Ready'

  onBroadcast?: (msg: SyncMessage) => void
  isApplyingRemote = false
  private transactOps: StoreOp[] | null = null

  private broadcastOp(op: StoreOp): void {
    if (this.isApplyingRemote) return
    if (this.transactOps !== null) {
      this.transactOps.push(op)
    } else {
      this.onBroadcast?.({ type: 'op', op })
    }
  }

  applyRemoteSync(msg: SyncMessage): void {
    this.isApplyingRemote = true
    try {
      applySyncMessage(this, msg)
    } finally {
      this.isApplyingRemote = false
    }
  }

  private listeners = new Set<() => void>()
  private undoStack: UndoEntry[] = []
  private redoStack: UndoEntry[] = []
  private group: UndoEntry[] | null = null
  private batchDepth = 0

  /** HyperFormula batches can't nest, so nested calls just run inline. */
  batch(fn: () => void): void {
    if (this.batchDepth > 0) return fn()
    this.batchDepth++
    try {
      this.hf.batch(fn)
    } finally {
      this.batchDepth--
    }
  }

  constructor() {
    this.hf = HyperFormula.buildEmpty(HF_CONFIG, BOOLEAN_NAMES)
    this.addSheetRaw('Sheet1', 0)
  }

  // ---------- subscription ----------
  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn)
    return () => this.listeners.delete(fn)
  }
  getVersion = (): number => this.version
  notify(): void {
    this.version++
    this.listeners.forEach((l) => l())
  }

  // ---------- sheet helpers ----------
  get active(): SheetState {
    return this.sheets[this.activeIndex]
  }
  hfId(sheet: SheetState): number {
    return this.hf.getSheetId(sheet.name)!
  }
  sheetByUid(uid: number): SheetState | undefined {
    return this.sheets.find((s) => s.uid === uid)
  }

  // ---------- reads ----------
  /** Text the formula bar shows / the editor starts with. */
  getRaw(sheet: SheetState, r: number, c: number): string {
    const addr = { sheet: this.hfId(sheet), row: r, col: c }
    const f = this.hf.getCellFormula(addr)
    if (f !== undefined) return f
    const v = this.hf.getCellSerialized(addr)
    if (v === null || v === undefined) return ''
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
    return String(v)
  }

  getValue(sheet: SheetState, r: number, c: number): CellValue {
    const v = this.hf.getCellValue({ sheet: this.hfId(sheet), row: r, col: c })
    if (v instanceof DetailedCellError) return { error: (v.value as never) ?? '#ERROR!' }
    return v as CellValue
  }

  getStyle(sheet: SheetState, r: number, c: number): CellStyle | undefined {
    return sheet.styles.get(cellKey(r, c))
  }

  getDisplay(sheet: SheetState, r: number, c: number): FormattedValue {
    return formatValue(this.getValue(sheet, r, c), this.getStyle(sheet, r, c)?.numFmt)
  }

  colWidth(sheet: SheetState, c: number): number {
    if (sheet.layout.hiddenCols[c]) return 0
    return sheet.layout.colWidths[c] ?? DEFAULT_COL_WIDTH
  }
  rowHeight(sheet: SheetState, r: number): number {
    if (sheet.layout.hiddenRows[r]) return 0
    return sheet.layout.rowHeights[r] ?? DEFAULT_ROW_HEIGHT
  }

  /** Bounding box of all cells with content or style, or null for an empty sheet. */
  usedRange(sheet: SheetState): Range | null {
    const d = this.hf.getSheetDimensions(this.hfId(sheet))
    let r2 = d.height - 1, c2 = d.width - 1
    for (const k of sheet.styles.keys()) {
      const [r, c] = k.split(',').map(Number)
      if (r > r2) r2 = r
      if (c > c2) c2 = c
    }
    return r2 < 0 || c2 < 0 ? null : { r1: 0, c1: 0, r2, c2 }
  }

  /** The merge range covering (r,c), if any. */
  mergeAt(sheet: SheetState, r: number, c: number): Range | undefined {
    return sheet.layout.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2)
  }

  /** The contiguous block of non-empty cells containing (r,c) — Excel's "current region",
   *  used to guess the table AutoFilter/Ctrl+T should cover. */
  currentRegion(sheet: SheetState, r: number, c: number): Range {
    const has = (rr: number, cc: number): boolean => rr >= 0 && cc >= 0 && this.getRaw(sheet, rr, cc) !== ''
    if (!has(r, c)) return { r1: r, c1: c, r2: r, c2: c }
    let r1 = r, r2 = r, c1 = c, c2 = c
    const rowHas = (rr: number): boolean => { if (rr < 0 || rr >= MAX_ROWS) return false; for (let cc = c1; cc <= c2; cc++) if (has(rr, cc)) return true; return false }
    const colHas = (cc: number): boolean => { if (cc < 0 || cc >= MAX_COLS) return false; for (let rr = r1; rr <= r2; rr++) if (has(rr, cc)) return true; return false }
    // Grow rows as far as they go, then columns, and repeat until neither side moves. Growing a
    // whole axis per pass keeps a 100k-row table linear instead of re-scanning every column
    // edge per added row. Bounded only by the sheet itself.
    for (;;) {
      let grew = false
      while (rowHas(r1 - 1)) { r1--; grew = true }
      while (rowHas(r2 + 1)) { r2++; grew = true }
      while (colHas(c1 - 1)) { c1--; grew = true }
      while (colHas(c2 + 1)) { c2++; grew = true }
      if (!grew) break
    }
    return { r1, c1, r2, c2 }
  }

  // ---------- AutoFilter ----------
  /** Recomputes which rows the current AutoFilter criteria hide, updating hiddenRows in place.
   *  Called from inside updateLayout, so it mutates the layout draft it's given directly. */
  private recomputeAutoFilter(sheet: SheetState, l: SheetLayout): void {
    const af = l.autoFilter
    if (!af) return
    for (const r of af.filteredOutRows) delete l.hiddenRows[r]
    const getDisplay = (r: number, c: number) => this.getDisplay(sheet, r, c).text
    const { hidden } = computeFilteredRows(af.range, af.columns, getDisplay)
    // Only rows this filter hides itself are recorded, so a row the user hid by hand stays
    // hidden when the filter is cleared or removed.
    af.filteredOutRows = hidden.filter((r) => !l.hiddenRows[r])
    for (const r of af.filteredOutRows) l.hiddenRows[r] = true
  }

  /** Data > Filter / Ctrl+Shift+L: toggles an AutoFilter over the current region. */
  toggleAutoFilter(sheet: SheetState = this.active): void {
    const on = !sheet.layout.autoFilter
    this.updateLayout(sheet, on ? 'AutoFilter' : 'Remove Filter', (l) => {
      if (!on) {
        for (const r of l.autoFilter?.filteredOutRows ?? []) delete l.hiddenRows[r]
        delete l.autoFilter
        return
      }
      const active = sheet.selection.active
      const sel = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
      const region = sel.r1 !== sel.r2 || sel.c1 !== sel.c2 ? sel : this.currentRegion(sheet, active.r, active.c)
      l.autoFilter = { range: region, columns: {}, filteredOutRows: [] }
    })
  }

  /** Sets (or, with `null`, clears) one column's checked-values criteria. */
  setColumnFilter(sheet: SheetState, col: number, criteria: FilterCriteria | null): void {
    this.updateLayout(sheet, criteria ? 'Filter' : 'Clear Filter', (l) => {
      if (!l.autoFilter) return
      if (criteria) l.autoFilter.columns[col] = criteria
      else delete l.autoFilter.columns[col]
      this.recomputeAutoFilter(sheet, l)
    })
  }

  /** Data > Clear: removes every column's criteria but keeps the AutoFilter dropdowns. */
  clearAllFilters(sheet: SheetState = this.active): void {
    this.updateLayout(sheet, 'Clear Filter', (l) => {
      if (!l.autoFilter) return
      l.autoFilter.columns = {}
      this.recomputeAutoFilter(sheet, l)
    })
  }

  /** Data > Reapply: re-runs the current criteria against (possibly changed) data. */
  reapplyFilter(sheet: SheetState = this.active): void {
    this.updateLayout(sheet, 'Reapply', (l) => this.recomputeAutoFilter(sheet, l))
  }

  /** Sorts every row of the AutoFilter range (keeping rows intact) by one column. */
  sortFilterRange(sheet: SheetState, col: number, order: 'asc' | 'desc'): void {
    const af = sheet.layout.autoFilter
    if (!af) return
    const n = af.range
    const rows: { r: number; cells: { raw: string; style?: CellStyle }[] }[] = []
    for (let r = n.r1 + 1; r <= n.r2; r++) {
      const cells: { raw: string; style?: CellStyle }[] = []
      for (let c = n.c1; c <= n.c2; c++) cells.push({ raw: this.getRaw(sheet, r, c), style: this.getStyle(sheet, r, c) })
      rows.push({ r, cells })
    }
    const key = (row: (typeof rows)[number]): [number, number | string] => {
      const v = this.getValue(sheet, row.r, col)
      if (v === null || v === undefined || v === '') return [2, '']
      if (typeof v === 'number') return [0, v]
      return [1, String(v).toLowerCase()]
    }
    const sorted = [...rows].sort((a, b) => {
      const [ta, va] = key(a)
      const [tb, vb] = key(b)
      if (ta !== tb) return ta - tb
      const cmp = va < vb ? -1 : va > vb ? 1 : 0
      return order === 'asc' ? cmp : -cmp
    })
    const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
    // A moved formula keeps pointing at the same relative cells (=B5*2 moving up 3 rows becomes
    // =B2*2), matching sortSelection in grid/commands.ts.
    sorted.forEach((row, i) => row.cells.forEach((cell, j) => {
      const r = n.r1 + 1 + i
      cells.push({ r, c: n.c1 + j, raw: shiftFormula(cell.raw, r - row.r, 0), style: cell.style })
    }))
    const label = order === 'asc' ? 'Sort A to Z' : 'Sort Z to A'
    // Rows move, so which ones the filter hides must be recomputed in the same undo step.
    this.transact(label, () => {
      this.setCells(sheet, cells, label)
      this.updateLayout(sheet, label, (l) => this.recomputeAutoFilter(sheet, l))
    })
  }

  // ---------- Notes & Comments ----------
  setNote(sheet: SheetState, r: number, c: number, text: string, author = DEFAULT_AUTHOR): void {
    const key = cellKey(r, c)
    this.updateLayout(sheet, text.trim() ? 'Note' : 'Delete Note', (l) => {
      if (text.trim() === '') delete l.notes[key]
      else l.notes[key] = { text, author }
    })
  }
  deleteNote(sheet: SheetState, r: number, c: number): void {
    this.setNote(sheet, r, c, '')
  }

  addComment(sheet: SheetState, r: number, c: number, text: string, author = DEFAULT_AUTHOR): void {
    if (!text.trim()) return
    const key = cellKey(r, c)
    this.updateLayout(sheet, 'New Comment', (l) => {
      l.comments[key] = { author, text, time: new Date().toISOString(), replies: [] }
    })
  }
  replyToComment(sheet: SheetState, r: number, c: number, text: string, author = DEFAULT_AUTHOR): void {
    if (!text.trim()) return
    const key = cellKey(r, c)
    this.updateLayout(sheet, 'Reply', (l) => {
      const t = l.comments[key]
      if (t) t.replies.push({ id: 'c' + Date.now() + Math.random().toString(36).slice(2, 7), author, text, time: new Date().toISOString() })
    })
  }
  resolveComment(sheet: SheetState, r: number, c: number, resolved = true): void {
    const key = cellKey(r, c)
    this.updateLayout(sheet, resolved ? 'Resolve Thread' : 'Reopen Thread', (l) => {
      const t = l.comments[key]
      if (t) t.resolved = resolved
    })
  }
  deleteComment(sheet: SheetState, r: number, c: number): void {
    const key = cellKey(r, c)
    this.updateLayout(sheet, 'Delete Comment', (l) => { delete l.comments[key] })
  }

  /** Review > Notes > Convert to Comments: every note becomes a (single-post) comment thread. */
  convertNotesToComments(sheet: SheetState = this.active): void {
    this.updateLayout(sheet, 'Convert to Comments', (l) => {
      for (const [key, note] of Object.entries(l.notes)) {
        l.comments[key] = { author: note.author, text: note.text, time: new Date().toISOString(), replies: [] }
      }
      l.notes = {}
    })
  }

  // ---------- undo machinery ----------
  /** Records an undoable change. `apply` is run immediately and again on redo. */
  commit(label: string, apply: () => void, revert: () => void): void {
    apply()
    const entry = { label, undo: revert, redo: apply }
    if (this.group) this.group.push(entry)
    else {
      this.undoStack.push(entry)
      if (this.undoStack.length > 100) this.undoStack.shift()
      this.redoStack = []
    }
    this.dirty = true
    this.notify()
  }

  /** Groups every commit made inside `fn` into a single undo step. */
  transact(label: string, fn: () => void): void {
    if (this.group) return fn()
    const g: UndoEntry[] = []
    this.group = g
    const isTopTransact = this.transactOps === null
    if (isTopTransact) this.transactOps = []
    try {
      fn() // not batched: fn may read cells between its commits
    } catch (e) {
      if (isTopTransact) this.transactOps = null // don't swallow every later broadcast
      throw e
    } finally {
      this.group = null
    }
    if (!g.length) {
      if (isTopTransact) this.transactOps = null
      return
    }
    this.undoStack.push({
      label,
      undo: () => this.batch(() => [...g].reverse().forEach((e) => e.undo())),
      redo: () => this.batch(() => g.forEach((e) => e.redo()))
    })
    this.redoStack = []
    if (isTopTransact) {
      const ops = this.transactOps ?? []
      this.transactOps = null
      if (ops.length && !this.isApplyingRemote) {
        this.onBroadcast?.({ type: 'transact', label, ops })
      }
    }
    this.notify()
  }

  undo(): void {
    const e = this.undoStack.pop()
    if (!e) return
    e.undo()
    this.redoStack.push(e)
    this.dirty = true
    if (!this.isApplyingRemote) {
      this.onBroadcast?.({ type: 'undo' })
    }
    this.notify()
  }
  redo(): void {
    const e = this.redoStack.pop()
    if (!e) return
    e.redo()
    this.undoStack.push(e)
    this.dirty = true
    if (!this.isApplyingRemote) {
      this.onBroadcast?.({ type: 'redo' })
    }
    this.notify()
  }
  get canUndo(): boolean { return this.undoStack.length > 0 }
  get canRedo(): boolean { return this.redoStack.length > 0 }
  get undoLabels(): string[] { return this.undoStack.map((e) => e.label).reverse() }
  get redoLabels(): string[] { return this.redoStack.map((e) => e.label).reverse() }

  // ---------- cell contents ----------
  private writeRaw(sheet: SheetState, r: number, c: number, raw: string): void {
    this.hf.setCellContents({ sheet: this.hfId(sheet), row: r, col: c }, raw === '' ? null : raw)
  }
  /** Writes many cells. HyperFormula's per-call bookkeeping inside a batch grows with the batch, so
   *  thousands of single-cell calls go quadratic; a dense block is written as one 2-D call instead. */
  private writeRaws(sheet: SheetState, list: { r: number; c: number; raw: string }[]): void {
    if (list.length < 64) { list.forEach((x) => this.writeRaw(sheet, x.r, x.c, x.raw)); return }
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1
    for (const x of list) { if (x.r < r1) r1 = x.r; if (x.r > r2) r2 = x.r; if (x.c < c1) c1 = x.c; if (x.c > c2) c2 = x.c }
    const h = r2 - r1 + 1, w = c2 - c1 + 1
    if (h * w > list.length * 2) {
      // Sparse: fall back to one call per row run.
      const byRow = new Map<number, { c: number; raw: string }[]>()
      for (const x of list) { let a = byRow.get(x.r); if (!a) byRow.set(x.r, (a = [])); a.push(x) }
      for (const [r, arr] of byRow) arr.forEach((x) => this.writeRaw(sheet, r, x.c, x.raw))
      return
    }
    // Dense: cells not in the list keep their current content.
    const id = this.hfId(sheet)
    const grid: (string | null)[][] = []
    const set = new Array<boolean>(h * w).fill(false)
    for (const x of list) set[(x.r - r1) * w + (x.c - c1)] = true
    for (let i = 0; i < h; i++) {
      const row: (string | null)[] = new Array(w)
      for (let j = 0; j < w; j++) row[j] = set[i * w + j] ? null : this.getRaw(sheet, r1 + i, c1 + j) || null
      grid.push(row)
    }
    for (const x of list) grid[x.r - r1][x.c - c1] = x.raw === '' ? null : x.raw
    this.hf.setCellContents({ sheet: id, row: r1, col: c1 }, grid)
  }
  private writeStyle(sheet: SheetState, r: number, c: number, style: CellStyle | undefined): void {
    const k = cellKey(r, c)
    if (style && Object.keys(style).length) sheet.styles.set(k, style)
    else sheet.styles.delete(k)
  }

  /** Sets a block of cells starting at (r,c). Rows of strings, as typed by the user. */
  setRangeRaw(sheet: SheetState, r: number, c: number, raws: string[][], label = 'Typing'): void {
    const before: { r: number; c: number; raw: string; style?: CellStyle }[] = []
    const after: { r: number; c: number; raw: string; style?: CellStyle }[] = []
    raws.forEach((row, i) =>
      row.forEach((raw, j) => {
        const rr = r + i, cc = c + j
        const oldStyle = this.getStyle(sheet, rr, cc)
        before.push({ r: rr, c: cc, raw: this.getRaw(sheet, rr, cc), style: oldStyle })
        let style = oldStyle
        const fmt = detectInputFormat(raw)
        if (fmt && (!oldStyle?.numFmt || oldStyle.numFmt === 'General')) style = { ...oldStyle, numFmt: fmt }
        after.push({ r: rr, c: cc, raw, style })
      })
    )
    const put = (list: typeof before) =>
      this.batch(() => {
        this.writeRaws(sheet, list)
        list.forEach((x) => this.writeStyle(sheet, x.r, x.c, x.style))
      })
    this.commit(label, () => put(after), () => put(before))
    this.broadcastOp({ type: 'setRangeRaw', sheetName: sheet.name, r, c, raws, label })
  }

  setRaw(sheet: SheetState, r: number, c: number, raw: string): void {
    this.setRangeRaw(sheet, r, c, [[raw]])
  }

  /** Writes full cell records (raw + style) — used by paste, fill and sort. */
  setCells(sheet: SheetState, cells: { r: number; c: number; raw: string; style?: CellStyle }[], label: string): void {
    const before = cells.map((x) => ({ r: x.r, c: x.c, raw: this.getRaw(sheet, x.r, x.c), style: this.getStyle(sheet, x.r, x.c) }))
    const put = (list: typeof cells) =>
      this.batch(() => {
        this.writeRaws(sheet, list)
        list.forEach((x) => this.writeStyle(sheet, x.r, x.c, x.style))
      })
    this.commit(label, () => put(cells), () => put(before))
    this.broadcastOp({ type: 'setCells', sheetName: sheet.name, cells, label })
  }

  // ---------- formatting ----------
  /** Merges `patch` into the style of every cell in `ranges`. Use `undefined` values to clear keys. */
  applyStyle(sheet: SheetState, ranges: Range[], patch: Partial<CellStyle> | ((s: CellStyle, r: number, c: number) => CellStyle), label = 'Format Cells'): void {
    const before = new Map<string, CellStyle | undefined>()
    const after = new Map<string, CellStyle | undefined>()
    for (const rg of ranges) {
      const n = this.clampToUsed(sheet, normRange(rg))
      for (let r = n.r1; r <= n.r2; r++)
        for (let c = n.c1; c <= n.c2; c++) {
          const k = cellKey(r, c)
          if (before.has(k)) continue
          const old = sheet.styles.get(k)
          before.set(k, old)
          const next = typeof patch === 'function' ? patch({ ...old }, r, c) : { ...old, ...patch }
          for (const key of Object.keys(next) as (keyof CellStyle)[]) if (next[key] === undefined) delete next[key]
          after.set(k, next)
        }
    }
    const put = (m: Map<string, CellStyle | undefined>) =>
      m.forEach((s, k) => {
        const [r, c] = k.split(',').map(Number)
        this.writeStyle(sheet, r, c, s)
      })
    this.commit(label, () => put(after), () => put(before))
    if (typeof patch === 'function') {
      const styles: Record<string, CellStyle | undefined> = {}
      after.forEach((s, k) => { styles[k] = s })
      this.broadcastOp({ type: 'setStyles', sheetName: sheet.name, styles, label })
    } else {
      this.broadcastOp({ type: 'applyStyle', sheetName: sheet.name, ranges, patch, label })
    }
  }

  /** Whole-row/column selections are limited to the used range so styling them stays cheap. */
  private clampToUsed(sheet: SheetState, rg: Range): Range {
    const used = this.usedRange(sheet)
    const lim = { r: Math.max(used?.r2 ?? 0, 200), c: Math.max(used?.c2 ?? 0, 50) }
    return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, rg.r1 + lim.r), c2: Math.min(rg.c2, rg.c1 + lim.c) }
  }

  clear(sheet: SheetState, ranges: Range[], what: 'all' | 'contents' | 'formats' = 'contents'): void {
    this.transact(what === 'formats' ? 'Clear Formats' : 'Clear', () => {
      for (const rg of ranges) {
        const n = this.clampToUsed(sheet, normRange(rg))
        if (what !== 'formats') {
          const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
          for (let r = n.r1; r <= n.r2; r++)
            for (let c = n.c1; c <= n.c2; c++)
              cells.push({ r, c, raw: '', style: what === 'all' ? undefined : this.getStyle(sheet, r, c) })
          this.setCells(sheet, cells, 'Clear')
        } else this.applyStyle(sheet, [n], () => ({}))
      }
    })
  }

  // ---------- layout ----------
  /** Undoable change to the sheet layout (widths, merges, freeze panes, gridlines, zoom, ...). */
  updateLayout(sheet: SheetState, label: string, fn: (l: SheetLayout) => void): void {
    const before = structuredClone(sheet.layout)
    const next = structuredClone(sheet.layout)
    fn(next)
    this.commit(label, () => (sheet.layout = structuredClone(next)), () => (sheet.layout = structuredClone(before)))
    if (label !== 'Zoom' && label !== 'Freeze Panes') {
      this.broadcastOp({ type: 'updateLayout', sheetName: sheet.name, label, layout: structuredClone(next) })
    }
  }

  setColWidth(sheet: SheetState, cols: number[], px: number): void {
    this.updateLayout(sheet, 'Column Width', (l) => cols.forEach((c) => (l.colWidths[c] = Math.max(0, px))))
  }
  setRowHeight(sheet: SheetState, rows: number[], px: number): void {
    this.updateLayout(sheet, 'Row Height', (l) => rows.forEach((r) => (l.rowHeights[r] = Math.max(0, px))))
  }

  // ---------- structure (insert/delete rows & columns) ----------
  /** Snapshot-based undo: captures the sheet's contents, styles and layout before and after. */
  private structural(sheet: SheetState, label: string, fn: () => void): void {
    const snap = () => ({
      content: this.hf.getSheetSerialized(this.hfId(sheet)),
      styles: new Map(sheet.styles),
      layout: structuredClone(sheet.layout)
    })
    const before = snap()
    fn()
    const after = snap()
    const restore = (s: typeof before) => {
      this.hf.setSheetContent(this.hfId(sheet), s.content)
      sheet.styles = new Map(s.styles)
      sheet.layout = structuredClone(s.layout)
    }
    this.commit(label, () => restore(after), () => restore(before))
  }

  private shiftKeys<T>(map: Record<number, T>, at: number, delta: number): Record<number, T> {
    const out: Record<number, T> = {}
    for (const [k, v] of Object.entries(map)) {
      const i = +k
      if (i < at) out[i] = v
      else if (delta > 0) out[i + delta] = v
      else if (i >= at - delta) out[i + delta] = v
    }
    return out
  }

  private shiftStyles(sheet: SheetState, axis: 'r' | 'c', at: number, delta: number): void {
    const next = new Map<string, CellStyle>()
    sheet.styles.forEach((s, k) => {
      let [r, c] = k.split(',').map(Number)
      const i = axis === 'r' ? r : c
      if (i >= at) {
        if (delta < 0 && i < at - delta) return
        if (axis === 'r') r += delta
        else c += delta
      }
      next.set(cellKey(r, c), s)
    })
    sheet.styles = next
  }

  private shiftMerges(sheet: SheetState, axis: 'r' | 'c', at: number, delta: number): void {
    const [a, b] = axis === 'r' ? (['r1', 'r2'] as const) : (['c1', 'c2'] as const)
    sheet.layout.merges = sheet.layout.merges
      .filter((m) => !(delta < 0 && m[a] >= at && m[b] < at - delta))
      .map((m) => {
        const n = { ...m }
        if (n[a] >= at) n[a] = Math.max(at, n[a] + delta)
        if (n[b] >= at) n[b] = Math.max(at, n[b] + delta)
        return n
      })
  }

  insertRows(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Insert Rows', () => {
      this.hf.addRows(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'r', at, count)
      this.shiftMerges(sheet, 'r', at, count)
      sheet.layout.rowHeights = this.shiftKeys(sheet.layout.rowHeights, at, count)
      sheet.layout.hiddenRows = this.shiftKeys(sheet.layout.hiddenRows, at, count)
    })
    this.broadcastOp({ type: 'insertRows', sheetName: sheet.name, at, count })
  }
  deleteRows(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Delete Rows', () => {
      this.hf.removeRows(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'r', at, -count)
      this.shiftMerges(sheet, 'r', at, -count)
      sheet.layout.rowHeights = this.shiftKeys(sheet.layout.rowHeights, at, -count)
      sheet.layout.hiddenRows = this.shiftKeys(sheet.layout.hiddenRows, at, -count)
    })
    this.broadcastOp({ type: 'deleteRows', sheetName: sheet.name, at, count })
  }
  insertCols(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Insert Columns', () => {
      this.hf.addColumns(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'c', at, count)
      this.shiftMerges(sheet, 'c', at, count)
      sheet.layout.colWidths = this.shiftKeys(sheet.layout.colWidths, at, count)
      sheet.layout.hiddenCols = this.shiftKeys(sheet.layout.hiddenCols, at, count)
    })
    this.broadcastOp({ type: 'insertCols', sheetName: sheet.name, at, count })
  }
  deleteCols(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Delete Columns', () => {
      this.hf.removeColumns(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'c', at, -count)
      this.shiftMerges(sheet, 'c', at, -count)
      sheet.layout.colWidths = this.shiftKeys(sheet.layout.colWidths, at, -count)
      sheet.layout.hiddenCols = this.shiftKeys(sheet.layout.hiddenCols, at, -count)
    })
    this.broadcastOp({ type: 'deleteCols', sheetName: sheet.name, at, count })
  }

  // ---------- sheets ----------
  private addSheetRaw(name: string, index: number): SheetState {
    this.hf.addSheet(name)
    const s: SheetState = { uid: nextUid++, name, styles: new Map(), layout: defaultLayout(), selection: defaultSelection() }
    this.sheets.splice(index, 0, s)
    return s
  }

  /** Next free 'SheetN' name. */
  newSheetName(): string {
    let n = this.sheets.length + 1
    const names = new Set(this.sheets.map((s) => s.name.toLowerCase()))
    while (names.has('sheet' + n)) n++
    return 'Sheet' + n
  }

  addSheet(name = this.newSheetName(), index = this.activeIndex + 1): void {
    let s: SheetState | undefined
    const prevActive = this.activeIndex
    this.commit(
      'Insert Sheet',
      () => {
        s = this.addSheetRaw(name, index)
        this.activeIndex = index
      },
      () => {
        this.hf.removeSheet(this.hfId(s!))
        this.sheets.splice(this.sheets.indexOf(s!), 1)
        this.activeIndex = prevActive
      }
    )
    this.broadcastOp({ type: 'addSheet', name, index })
  }

  deleteSheet(index: number): void {
    if (this.sheets.length <= 1) return
    const s = this.sheets[index]
    const content = this.hf.getSheetSerialized(this.hfId(s))
    const prevActive = this.activeIndex
    this.commit(
      'Delete Sheet',
      () => {
        this.hf.removeSheet(this.hfId(s))
        this.sheets.splice(index, 1)
        this.activeIndex = Math.min(index, this.sheets.length - 1)
      },
      () => {
        this.hf.addSheet(s.name)
        this.hf.setSheetContent(this.hfId(s), content)
        this.sheets.splice(index, 0, s)
        this.activeIndex = prevActive
      }
    )
    this.broadcastOp({ type: 'deleteSheet', index })
  }

  /** Returns an error message (Excel's wording) or null on success. */
  renameSheet(index: number, name: string): string | null {
    const s = this.sheets[index]
    const trimmed = name.trim()
    if (!trimmed || trimmed.length > 31 || /[\\/?*[\]:]/.test(trimmed) || /^'|'$/.test(trimmed))
      return "You typed an invalid name for a sheet or chart. Make sure that:\n\n• The name that you type does not exceed 31 characters.\n• The name does not contain any of the following characters: \\ / ? * [ ] :\n• You did not leave the name blank."
    if (trimmed.toLowerCase() === 'history') return "'History' is a reserved name. Please choose a different name."
    if (this.sheets.some((o, i) => i !== index && o.name.toLowerCase() === trimmed.toLowerCase()))
      return 'That name is already taken. Try a different one.'
    const old = s.name
    if (old === trimmed) return null
    this.commit(
      'Rename Sheet',
      () => { this.hf.renameSheet(this.hf.getSheetId(old)!, trimmed); s.name = trimmed },
      () => { this.hf.renameSheet(this.hf.getSheetId(trimmed)!, old); s.name = old }
    )
    this.broadcastOp({ type: 'renameSheet', index, name: trimmed })
    return null
  }

  moveSheet(from: number, to: number): void {
    if (from === to) return
    const move = (a: number, b: number) => {
      const active = this.sheets[this.activeIndex]
      const [s] = this.sheets.splice(a, 1)
      this.sheets.splice(b, 0, s)
      this.activeIndex = this.sheets.indexOf(active)
    }
    this.commit('Move Sheet', () => move(from, to), () => move(to, from))
    this.broadcastOp({ type: 'moveSheet', from, to })
  }

  setActiveSheet(index: number): void {
    if (index < 0 || index >= this.sheets.length || index === this.activeIndex) return
    this.activeIndex = index
    this.clipboard = null
    this.notify()
  }

  // ---------- selection (not undoable) ----------
  setSelection(sel: Selection, sheet: SheetState = this.active): void {
    sheet.selection = sel
    this.notify()
  }

  setMode(mode: WorkbookStore['mode']): void {
    if (this.mode !== mode) {
      this.mode = mode
      this.notify()
    }
  }

  // ---------- whole workbook ----------
  load(data: WorkbookData, fileName: string, filePath: string | null): void {
    // Build HyperFormula in one pass from dense 2-D arrays: setting cells one by one
    // (even inside a batch) is orders of magnitude slower on large files.
    const contents: Record<string, (string | null)[][]> = {}
    this.sheets = []
    data.sheets.forEach((sd) => {
      const rows: (string | null)[][] = []
      const s: SheetState = { uid: nextUid++, name: sd.name, styles: new Map(), layout: { ...defaultLayout(), ...sd.layout }, selection: defaultSelection() }
      for (const k in sd.cells) {
        const cell = sd.cells[k]
        const comma = k.indexOf(',')
        const r = +k.slice(0, comma), c = +k.slice(comma + 1)
        if (cell.raw) (rows[r] ??= [])[c] = cell.raw
        if (cell.style && Object.keys(cell.style).length) s.styles.set(k, cell.style)
      }
      // HyperFormula wants no holes in the outer array.
      for (let r = 0; r < rows.length; r++) rows[r] ??= []
      contents[sd.name] = rows
      this.sheets.push(s)
    })
    this.hf.destroy()
    this.hf = data.sheets.length ? HyperFormula.buildFromSheets(contents, HF_CONFIG, BOOLEAN_NAMES) : HyperFormula.buildEmpty(HF_CONFIG, BOOLEAN_NAMES)
    if (!this.sheets.length) this.addSheetRaw('Sheet1', 0)
    this.activeIndex = Math.min(data.activeSheet ?? 0, this.sheets.length - 1)
    this.fileName = fileName
    this.filePath = filePath
    this.dirty = false
    this.clipboard = null
    this.undoStack = []
    this.redoStack = []
    this.notify()
  }

  toData(): WorkbookData {
    return {
      activeSheet: this.activeIndex,
      sheets: this.sheets.map((s) => {
        const cells: WorkbookData['sheets'][number]['cells'] = {}
        const id = this.hfId(s)
        const content = this.hf.getSheetSerialized(id)
        content.forEach((row, r) =>
          row.forEach((v, c) => {
            if (v === null || v === undefined || v === '') return
            cells[cellKey(r, c)] = { raw: typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v) }
          })
        )
        s.styles.forEach((style, k) => (cells[k] = { ...cells[k], style }))
        return { name: s.name, cells, layout: structuredClone(s.layout) }
      })
    }
  }

  newWorkbook(): void {
    this.load({ sheets: [{ name: 'Sheet1', cells: {} }] }, 'Book1', null)
  }
}

export const store = new WorkbookStore()

/** React hook: re-renders the component whenever the store changes. Returns the store. */
export function useStore(): WorkbookStore {
  useSyncExternalStore(store.subscribe, store.getVersion)
  return store
}

// Handy for debugging in DevTools.
;(globalThis as unknown as { store: WorkbookStore }).store = store

// The workbook store: single source of truth for the app.
// Cell contents and formula evaluation live in HyperFormula; styles, layout and selection live here.
// Every mutation goes through `commit`, which records an undo entry and notifies subscribers.
import { HyperFormula, DetailedCellError } from 'hyperformula'
import { useSyncExternalStore } from 'react'
import {
  CellStyle, CellValue, Range, Selection, SheetLayout, cellKey, normRange,
  DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, MAX_ROWS, MAX_COLS, FilterCriteria, WorkbookProtection, CalcMode, SheetProtection, ProtectAllowKey, NoteData, CommentThread, rangeA1
, normalFont, DEFAULT_FONT, DEFAULT_FONT_SIZE } from './types'
import { formatValue, detectInputFormat, numericInput, FormattedValue } from '../format/numfmt'
import { computeFilteredRows } from '../engine/filter'
import { shiftFormula } from '../engine/shiftFormula'
import { expandFormula } from '../engine/lambda'
import { eachPacked, type PackedCells } from '../io/packedCells'
import { patchHyperFormulaGraph, instrumentHf } from '../engine/hfPatches'
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
    /** `value`: a formula's last computed result, saved as Excel's cached value.
     *  `spill`: for a formula whose result spills (FILTER, UNIQUE...), the covered range, e.g. "H8:H10". */
    cells: Record<string, { raw?: string; style?: CellStyle; value?: CellValue; spill?: string }> // key 'r,c'
    /** Big sheets arrive from the parse worker in this compact form instead of `cells` (see io/packedCells.ts). */
    packed?: PackedCells
    layout?: Partial<SheetLayout>
  }[]
  activeSheet?: number
  /** Review > Protect Workbook (structure / windows lock). */
  protection?: WorkbookProtection
  /** Formulas > Calculation Options; 'auto' when absent. */
  calcMode?: CalcMode
  /** The workbook's Normal-style font when it isn't Calibri 11 (kept so a save doesn't reset it). */
  defaultFont?: { name: string; size: number }
  /** Workbook defined names (named ranges, print areas...). `ref` has no leading '='; `sheet` is
   *  the index of the sheet a local name belongs to (localSheetId). */
  names?: DefinedNameData[]
}

export interface DefinedNameData { name: string; ref: string; sheet?: number; hidden?: boolean }

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
import { shiftLayoutRefs, shiftRange } from './shiftLayout'
import { registerExtraFunctions, BOOLEAN_NAMES, setRowVisibility } from '../engine/extraFunctions'
import { DATE_FORMATS } from '../engine/locale'
registerExtraFunctions()
patchHyperFormulaGraph()

const HF_CONFIG = {
  maxRows: 1048576,
  maxColumns: 16384,
  licenseKey: 'gpl-v3',
  undoLimit: 0,
  useArrayArithmetic: true,
  dateFormats: DATE_FORMATS,
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  currencySymbol: ['$'],
  nullDate: { year: 1899, month: 12, day: 30 }
}

function sameStyle(a: CellStyle | undefined, b: CellStyle | undefined): boolean {
  if (a === b) return true
  const ea = !a || Object.keys(a).length === 0, eb = !b || Object.keys(b).length === 0
  if (ea || eb) return ea && eb
  return JSON.stringify(a) === JSON.stringify(b)
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
  /** Defined names carried through load/save; ordinary ones are also live in the engine. */
  names: (DefinedNameData & { sheetUid?: number })[] = []
  /** Review > Protect Workbook. */
  bookProtection: WorkbookProtection | undefined
  /** Formulas > Calculation Options. */
  calcMode: CalcMode = 'auto'
  defaultFont?: { name: string; size: number }
  /** Manual mode: formula results as of the last calculation. Excel leaves them stale until F9. */
  private frozen: Map<string, { f: string; v: CellValue }> | null = null
  /** Manual mode: something changed since the last calculation (the status bar shows "Calculate"). */
  calcPending = false
  /** The UI shows Excel's "protected sheet" message through this. */
  onProtectedEdit?: (msg: string) => void

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
    this.hf = instrumentHf(HyperFormula.buildEmpty(HF_CONFIG, BOOLEAN_NAMES))
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
    const raw = this.engineRaw(sheet, r, c)
    const e = this.letSrc.get(`${sheet.uid}:${r},${c}`)
    return e && e.expanded === raw ? e.orig : raw
  }

  /**
   * Reader for many cells at once: one engine call for the whole sheet instead of two per cell
   * (about 6x faster for sorts, pastes and clears over tens of thousands of cells).
   */
  rawReader(sheet: SheetState, count: number): (r: number, c: number) => string {
    if (count < 3000) return (r, c) => this.getRaw(sheet, r, c)
    const m = this.hf.getSheetSerialized(this.hfId(sheet)) as (string | number | boolean | null | undefined)[][]
    const lets = this.letSrc.size > 0
    return (r, c) => {
      const v = m[r]?.[c]
      if (v === null || v === undefined) return ''
      const raw = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)
      if (lets && raw.charCodeAt(0) === 61) { const e = this.letSrc.get(`${sheet.uid}:${r},${c}`); if (e && e.expanded === raw) return e.orig }
      return raw
    }
  }

  /** LET formulas as the engine holds them (expanded), keyed by cell. */
  private letSrc = new Map<string, { orig: string; expanded: string }>()
  private toEngine(sheet: SheetState, r: number, c: number, raw: string): string {
    const key = `${sheet.uid}:${r},${c}`
    const x = raw.charCodeAt(0) === 61 ? expandFormula(raw) : null
    if (x === null) { this.letSrc.delete(key); return raw }
    this.letSrc.set(key, { orig: raw, expanded: x })
    return x
  }
  private engineRaw(sheet: SheetState, r: number, c: number): string {
    const addr = { sheet: this.hfId(sheet), row: r, col: c }
    const f = this.hf.getCellFormula(addr)
    if (f !== undefined) return f
    const v = this.hf.getCellSerialized(addr)
    if (v === null || v === undefined) return ''
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
    return String(v)
  }

  /** True when the cell has no content (one engine call, unlike getRaw); used by Ctrl+Arrow jumps and region detection. */
  isBlank(sheet: SheetState, r: number, c: number): boolean {
    const v = this.hf.getCellSerialized({ sheet: this.hfId(sheet), row: r, col: c })
    return v === null || v === undefined || v === ''
  }

  getValue(sheet: SheetState, r: number, c: number): CellValue {
    const addr = { sheet: this.hfId(sheet), row: r, col: c }
    if (this.frozen) {
      const e = this.frozen.get(`${addr.sheet}:${r},${c}`)
      if (e && this.hf.getCellFormula(addr) === e.f) return e.v
    }
    const v = this.hf.getCellValue(addr)
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
    // The extent of the styled cells is cached: scanning every style key on each render made arrow keys take
    // 400 ms on a 50,000-row sheet. writeStyle drops the cache when a key is added or removed.
    let e = this.usedCache.get(sheet)
    if (!e || e.styles !== sheet.styles || e.size !== sheet.styles.size) {
      let sr = -1, sc = -1
      for (const k of sheet.styles.keys()) {
        const comma = k.indexOf(',')
        const r = +k.slice(0, comma), c = +k.slice(comma + 1)
        if (r > sr) sr = r
        if (c > sc) sc = c
      }
      e = { styles: sheet.styles, size: sheet.styles.size, r2: sr, c2: sc }
      this.usedCache.set(sheet, e)
    }
    const r2 = Math.max(d.height - 1, e.r2), c2 = Math.max(d.width - 1, e.c2)
    return r2 < 0 || c2 < 0 ? null : { r1: 0, c1: 0, r2, c2 }
  }
  private usedCache = new WeakMap<SheetState, { styles: Map<string, CellStyle>; size: number; r2: number; c2: number }>()

  /** The merge range covering (r,c), if any. */
  mergeAt(sheet: SheetState, r: number, c: number): Range | undefined {
    return sheet.layout.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2)
  }

  /** The contiguous block of non-empty cells containing (r,c) — Excel's "current region",
   *  used to guess the table AutoFilter/Ctrl+T should cover. */
  currentRegion(sheet: SheetState, r: number, c: number): Range {
    const has = (rr: number, cc: number): boolean => rr >= 0 && cc >= 0 && !this.isBlank(sheet, rr, cc)
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
    if (sheet.layout.protection) { this.deny(WorkbookStore.PROTECTED_MSG); return } // Excel can't add or remove a filter on a protected sheet
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
    if (sheet.layout.protection && !this.allows(sheet, 'autoFilter')) return
    this.updateLayout(sheet, criteria ? 'Filter' : 'Clear Filter', (l) => {
      if (!l.autoFilter) return
      if (criteria) l.autoFilter.columns[col] = criteria
      else delete l.autoFilter.columns[col]
      this.recomputeAutoFilter(sheet, l)
    })
  }

  /** Data > Clear: removes every column's criteria but keeps the AutoFilter dropdowns. */
  clearAllFilters(sheet: SheetState = this.active): void {
    if (sheet.layout.protection && !this.allows(sheet, 'autoFilter')) return
    this.updateLayout(sheet, 'Clear Filter', (l) => {
      if (l.advFilter) { for (const r of l.advFilter.hidden) delete l.hiddenRows[r]; delete l.advFilter }
      if (!l.autoFilter) return
      l.autoFilter.columns = {}
      this.recomputeAutoFilter(sheet, l)
    })
  }

  /** Data > Reapply: re-runs the current criteria against (possibly changed) data. */
  reapplyFilter(sheet: SheetState = this.active): void {
    if (sheet.layout.protection && !this.allows(sheet, 'autoFilter')) return
    this.updateLayout(sheet, 'Reapply', (l) => this.recomputeAutoFilter(sheet, l))
  }

  /** Sorts every row of the AutoFilter range (keeping rows intact) by one column. */
  sortFilterRange(sheet: SheetState, col: number, order: 'asc' | 'desc'): void {
    const af = sheet.layout.autoFilter
    if (!af) return
    if (sheet.layout.protection && (!this.allows(sheet, 'sort') || !this.canEditRange(sheet, af.range))) return
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
      this.moveRowAnnotations(sheet, new Map(sorted.map((row, i) => [row.r, n.r1 + 1 + i])), n.c1, n.c2, label)
      this.updateLayout(sheet, label, (l) => this.recomputeAutoFilter(sheet, l))
    })
  }

  /** After a sort moved rows (origRow -> newRow within columns c1..c2), moves the notes, comments
   *  and links in those cells along with them, as Excel does. */
  moveRowAnnotations(sheet: SheetState, moves: Map<number, number>, c1: number, c2: number, label: string): void {
    const has = (m?: Record<string, unknown>) => m && Object.keys(m).length
    if (!has(sheet.layout.notes) && !has(sheet.layout.comments) && !has(sheet.layout.hyperlinks)) return
    const remap = <T>(m: Record<string, T> | undefined): Record<string, T> | undefined => {
      if (!m) return m
      const out: Record<string, T> = {}
      const moved: [string, T][] = []
      for (const [k, v] of Object.entries(m)) {
        const [r, c] = k.split(',').map(Number)
        const to = c >= c1 && c <= c2 ? moves.get(r) : undefined
        if (to === undefined) out[k] = v
        else moved.push([`${to},${c}`, v])
      }
      for (const [k, v] of moved) out[k] = v
      return out
    }
    this.updateLayout(sheet, label, (l) => {
      l.notes = remap(l.notes) ?? {}
      l.comments = remap(l.comments) ?? {}
      if (l.hyperlinks) l.hyperlinks = remap(l.hyperlinks)
    })
  }

  // ---------- Notes & Comments ----------
  setNote(sheet: SheetState, r: number, c: number, text: string, author = DEFAULT_AUTHOR): void {
    if (sheet.layout.protection && !this.allows(sheet, 'objects')) return
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
    if (sheet.layout.protection && !this.allows(sheet, 'objects')) return
    if (!text.trim()) return
    const key = cellKey(r, c)
    this.updateLayout(sheet, 'New Comment', (l) => {
      l.comments[key] = { author, text, time: new Date().toISOString(), replies: [] }
    })
  }
  replyToComment(sheet: SheetState, r: number, c: number, text: string, author = DEFAULT_AUTHOR): void {
    if (sheet.layout.protection && !this.allows(sheet, 'objects')) return
    if (!text.trim()) return
    const key = cellKey(r, c)
    this.updateLayout(sheet, 'Reply', (l) => {
      const t = l.comments[key]
      if (t) t.replies.push({ id: 'c' + Date.now() + Math.random().toString(36).slice(2, 7), author, text, time: new Date().toISOString() })
    })
  }
  resolveComment(sheet: SheetState, r: number, c: number, resolved = true): void {
    if (sheet.layout.protection && !this.allows(sheet, 'objects')) return
    const key = cellKey(r, c)
    this.updateLayout(sheet, resolved ? 'Resolve Thread' : 'Reopen Thread', (l) => {
      const t = l.comments[key]
      if (t) t.resolved = resolved
    })
  }
  deleteComment(sheet: SheetState, r: number, c: number): void {
    if (sheet.layout.protection && !this.allows(sheet, 'objects')) return
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

  // ---------- protection ----------
  static readonly PROTECTED_MSG = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be prompted for a password."
  static readonly BOOK_PROTECTED_MSG = "The workbook is protected, so its structure can't be changed. To make a change, unprotect the workbook. You might be prompted for a password."
  private denying = false
  private deny(msg: string): false {
    if (!this.denying) {
      this.denying = true
      queueMicrotask(() => { this.denying = false })
      this.onProtectedEdit?.(msg)
    }
    return false
  }
  /** True when the sheet isn't protected or the protection allows `key`; otherwise tells the user and returns false. */
  allows(sheet: SheetState, key: ProtectAllowKey): boolean {
    const p = sheet.layout.protection
    return !p || p.allow[key] || this.deny(WorkbookStore.PROTECTED_MSG)
  }
  cellLocked(sheet: SheetState, r: number, c: number): boolean {
    return sheet.styles.get(cellKey(r, c))?.locked !== false
  }
  /** False (after telling the user) if any of these cells is locked on a protected sheet. */
  canEditCells(sheet: SheetState, cells: Iterable<{ r: number; c: number }>): boolean {
    if (!sheet.layout.protection) return true
    for (const x of cells) if (this.cellLocked(sheet, x.r, x.c)) return this.deny(WorkbookStore.PROTECTED_MSG)
    return true
  }
  canEditRange(sheet: SheetState, rg: Range): boolean {
    if (!sheet.layout.protection) return true
    const n = this.clampToUsed(sheet, normRange(rg))
    for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) if (this.cellLocked(sheet, r, c)) return this.deny(WorkbookStore.PROTECTED_MSG)
    return true
  }
  /** Sheet add/delete/rename/move/hide are refused while the workbook structure is locked. */
  /** For commands Excel disables outright on a protected sheet (Conditional Formatting, Data Validation...). */
  refuseIfProtected(sheet: SheetState): boolean {
    return !!sheet.layout.protection && !this.deny(WorkbookStore.PROTECTED_MSG)
  }
  canChangeStructure(): boolean {
    return !this.bookProtection?.lockStructure || this.deny(WorkbookStore.BOOK_PROTECTED_MSG)
  }
  /** Not undoable (like Excel): otherwise Undo could remove a password-protected sheet's protection. */
  setSheetProtection(sheet: SheetState, prot: SheetProtection | null): void {
    if (prot) sheet.layout = { ...sheet.layout, protection: prot }
    else { const l = { ...sheet.layout }; delete l.protection; sheet.layout = l }
    this.dirty = true
    this.broadcastOp({ type: 'updateLayout', sheetName: sheet.name, label: prot ? 'Protect Sheet' : 'Unprotect Sheet', layout: structuredClone(sheet.layout) })
    this.notify()
  }
  setBookProtection(p: WorkbookProtection | undefined): void {
    this.bookProtection = p && (p.lockStructure || p.lockWindows) ? p : undefined
    this.dirty = true
    this.notify()
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
    if (this.frozen) this.calcPending = true
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
    this.hf.setCellContents({ sheet: this.hfId(sheet), row: r, col: c }, raw === '' ? null : this.toEngine(sheet, r, c, raw))
  }
  /** Writes many cells. HyperFormula's per-call bookkeeping inside a batch grows with the batch, so
   *  thousands of single-cell calls go quadratic; a dense block is written as one 2-D call instead. */
  private writeRaws(sheet: SheetState, list: { r: number; c: number; raw: string }[]): void {
    if (list.length < 64) { list.forEach((x) => this.writeRaw(sheet, x.r, x.c, x.raw)); return }
    const id = this.hfId(sheet)
    // Runs of consecutive cells in a row, then runs of identical column spans in consecutive rows, become one 2-D write.
    const byRow = new Map<number, { c: number; raw: string }[]>()
    for (const x of list) { let a = byRow.get(x.r); if (!a) byRow.set(x.r, (a = [])); a.push(x) }
    interface Seg { r: number; c1: number; raws: string[] }
    const segs: Seg[] = []
    for (const r of [...byRow.keys()].sort((p, q) => p - q)) {
      const cells = byRow.get(r)!.sort((p, q) => p.c - q.c)
      let cur: Seg | null = null
      for (const x of cells) {
        if (cur && x.c === cur.c1 + cur.raws.length) cur.raws.push(x.raw)
        else { cur = { r, c1: x.c, raws: [x.raw] }; segs.push(cur) }
      }
    }
    for (let i = 0; i < segs.length; ) {
      let j = i + 1
      while (j < segs.length && segs[j].r === segs[j - 1].r + 1 && segs[j].c1 === segs[i].c1 && segs[j].raws.length === segs[i].raws.length) j++
      const block = segs.slice(i, j).map((sg) => sg.raws.map((raw, k) => (raw === '' ? null : this.toEngine(sheet, sg.r, sg.c1 + k, raw))))
      this.hf.setCellContents({ sheet: id, row: segs[i].r, col: segs[i].c1 }, block)
      i = j
    }
  }
  private writeStyle(sheet: SheetState, r: number, c: number, style: CellStyle | undefined): void {
    const k = cellKey(r, c)
    if (style && Object.keys(style).length) {
      if (!sheet.styles.has(k)) this.usedCache.delete(sheet)
      sheet.styles.set(k, style)
    } else if (sheet.styles.delete(k)) this.usedCache.delete(sheet)
  }

  /** Sets a block of cells starting at (r,c). Rows of strings, as typed by the user. */
  setRangeRaw(sheet: SheetState, r: number, c: number, raws: string[][], label = 'Typing'): void {
    if (sheet.layout.protection && !this.canEditRange(sheet, { r1: r, c1: c, r2: r + raws.length - 1, c2: c + Math.max(0, ...raws.map((x) => x.length)) - 1 })) return
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
        // A Text-formatted (@) cell keeps what was typed as text, like Excel: "00012" stays "00012".
        const asText = oldStyle?.numFmt === '@' && raw !== '' && !raw.startsWith("'") && !raw.startsWith('=')
        after.push({ r: rr, c: cc, raw: asText ? "'" + raw : numericInput(raw) ?? raw, style })
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
    if (sheet.layout.protection && !this.canEditCells(sheet, cells)) return
    const read = this.rawReader(sheet, cells.length)
    // Rows in which every cell already holds exactly this content and style are left alone (a sort, paste or fill often
    // rewrites unchanged rows, and every write costs the engine a dependency recalculation). Changed rows are written whole.
    const rowChanged = new Set<number>()
    const infos = cells.map((x) => {
      const raw = read(x.r, x.c)
      const style = this.getStyle(sheet, x.r, x.c)
      if (!(raw === x.raw && sameStyle(style, x.style))) rowChanged.add(x.r)
      return { r: x.r, c: x.c, raw, style }
    })
    if (!rowChanged.size) return
    const before: typeof cells = []
    const changed: typeof cells = []
    cells.forEach((x, i) => { if (rowChanged.has(x.r)) { before.push(infos[i]); changed.push(x) } })
    cells = changed
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
    if (sheet.layout.protection) {
      if (!this.allows(sheet, 'formatCells')) return
      if (typeof patch !== 'function' && ('locked' in patch || 'hidden' in patch)) { this.deny(WorkbookStore.PROTECTED_MSG); return }
    }
    // Cells are collected in arrays (not string-keyed maps) and, for a plain patch, cells that share one old style share
    // one new style object: formatting the whole 50,000-row sheet went from 4 s to well under a second.
    const cells: { r: number; c: number; old: CellStyle | undefined; next: CellStyle }[] = []
    const seen = new Set<number>()
    const shared = typeof patch === 'function' ? null : new Map<CellStyle | undefined, CellStyle>()
    for (const rg of ranges) {
      const n = this.clampToUsed(sheet, normRange(rg))
      for (let r = n.r1; r <= n.r2; r++)
        for (let c = n.c1; c <= n.c2; c++) {
          const id = r * MAX_COLS + c
          if (seen.has(id)) continue
          seen.add(id)
          const old = sheet.styles.get(cellKey(r, c))
          let next = shared?.get(old)
          if (!next) {
            next = typeof patch === 'function' ? patch({ ...old }, r, c) : { ...old, ...patch }
            for (const key of Object.keys(next) as (keyof CellStyle)[]) if (next[key] === undefined) delete next[key]
            shared?.set(old, next)
          }
          cells.push({ r, c, old, next })
        }
    }
    const put = (useNext: boolean): void => { for (const x of cells) this.writeStyle(sheet, x.r, x.c, useNext ? x.next : x.old) }
    this.commit(label, () => put(true), () => put(false))
    if (typeof patch === 'function') {
      if (this.onBroadcast) {
        const styles: Record<string, CellStyle | undefined> = {}
        for (const x of cells) styles[cellKey(x.r, x.c)] = x.next
        this.broadcastOp({ type: 'setStyles', sheetName: sheet.name, styles, label })
      }
    } else {
      this.broadcastOp({ type: 'applyStyle', sheetName: sheet.name, ranges, patch, label })
    }
  }

  /** Whole-row/column selections are limited to the used range so styling them stays cheap. */
  clampToUsed(sheet: SheetState, rg: Range): Range {
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
      if (what === 'all') {
        // Excel's Clear All also removes notes, comments and hyperlinks in the range.
        const rs = ranges.map(normRange)
        const inSel = (k: string) => { const [r, c] = k.split(',').map(Number); return rs.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) }
        const l = sheet.layout
        if (Object.keys(l.notes).some(inSel) || Object.keys(l.comments).some(inSel) || Object.keys(l.hyperlinks ?? {}).some(inSel) || (l.sparklines ?? []).some((z) => inSel(z.r + ',' + z.c))) {
          this.updateLayout(sheet, 'Clear', (x) => {
            x.notes = Object.fromEntries(Object.entries(x.notes).filter(([k]) => !inSel(k)))
            x.comments = Object.fromEntries(Object.entries(x.comments).filter(([k]) => !inSel(k)))
            if (x.hyperlinks) x.hyperlinks = Object.fromEntries(Object.entries(x.hyperlinks).filter(([k]) => !inSel(k)))
            if (x.sparklines) x.sparklines = x.sparklines.filter((z) => !inSel(z.r + ',' + z.c))
          })
        }
      }
    })
  }

  // ---------- layout ----------
  /** Undoable change to the sheet layout (widths, merges, freeze panes, gridlines, zoom, ...). */
  /** SUBTOTAL/AGGREGATE are volatile; any engine change recalculates them after rows are hidden/shown. */
  private visVersion = 0
  touchVisibility(): void {
    try { this.hf.changeNamedExpression('ZZROWVISVER', '=' + ++this.visVersion) } catch { /* not present */ }
  }
  private filteredCache = new WeakMap<number[], Set<number>>()
  rowHiddenKind(hfSheet: number, row: number): 0 | 1 | 2 {
    const sh = this.sheets.find((s) => this.hfId(s) === hfSheet)
    if (!sh) return 0
    const fr = sh.layout.autoFilter?.filteredOutRows
    if (fr && fr.length) {
      let set = this.filteredCache.get(fr)
      if (!set) this.filteredCache.set(fr, (set = new Set(fr)))
      if (set.has(row)) return 1
    }
    return sh.layout.hiddenRows?.[row] ? 2 : 0
  }

  /** `need`: the protection permission this change requires on a protected sheet ('blocked' = never allowed). */
  updateLayout(sheet: SheetState, label: string, fn: (l: SheetLayout) => void, need?: ProtectAllowKey | 'blocked'): void {
    if (need && sheet.layout.protection && (need === 'blocked' ? !this.deny(WorkbookStore.PROTECTED_MSG) : !this.allows(sheet, need))) return
    const before = structuredClone(sheet.layout)
    const next = structuredClone(sheet.layout)
    fn(next)
    const rowsChanged = before.autoFilter?.filteredOutRows !== next.autoFilter?.filteredOutRows || JSON.stringify(before.hiddenRows) !== JSON.stringify(next.hiddenRows) || JSON.stringify(before.autoFilter?.filteredOutRows) !== JSON.stringify(next.autoFilter?.filteredOutRows)
    // Protection is never rolled back by Undo/Redo (they restore whole-layout snapshots).
    const apply = (l: SheetLayout) => { const p = sheet.layout.protection; sheet.layout = structuredClone(l); if (p) sheet.layout.protection = p; else delete sheet.layout.protection; if (rowsChanged) this.touchVisibility() }
    this.commit(label, () => apply(next), () => apply(before))
    if (label !== 'Zoom' && label !== 'Freeze Panes') {
      this.broadcastOp({ type: 'updateLayout', sheetName: sheet.name, label, layout: structuredClone(next) })
    }
  }

  setColWidth(sheet: SheetState, cols: number[], px: number): void {
    this.updateLayout(sheet, 'Column Width', (l) => cols.forEach((c) => (l.colWidths[c] = Math.max(0, px))), 'formatColumns')
  }
  setRowHeight(sheet: SheetState, rows: number[], px: number): void {
    this.updateLayout(sheet, 'Row Height', (l) => rows.forEach((r) => (l.rowHeights[r] = Math.max(0, px))), 'formatRows')
  }

  // ---------- structure (insert/delete rows & columns) ----------
  /** Snapshot-based undo: captures the sheet's contents, styles and layout before and after. */
  /**
   * Insert/delete rows or columns. The engine op is applied directly and undone by its inverse (remove/add), which
   * keeps undo as fast as the edit; deleting also snapshots the sheet so undo can restore the deleted cells and any
   * formulas that lost their references (only the cells that differ are rewritten).
   */
  private structural(sheet: SheetState, label: string, op: { kind: 'insert' | 'delete'; dim: 'r' | 'c'; at: number; count: number }, shiftLayout: () => void): void {
    const id = this.hfId(sheet)
    const before = { styles: new Map(sheet.styles), layout: structuredClone(sheet.layout) }
    const content = op.kind === 'delete' ? (this.hf.getSheetSerialized(id) as (string | number | boolean | null)[][]) : null
    const engine = (add: boolean): void => {
      const span: [number, number] = [op.at, op.count]
      if (op.dim === 'r') add ? this.hf.addRows(id, span) : this.hf.removeRows(id, span)
      else add ? this.hf.addColumns(id, span) : this.hf.removeColumns(id, span)
    }
    engine(op.kind === 'insert')
    shiftLayout()
    const after = { styles: new Map(sheet.styles), layout: structuredClone(sheet.layout) }
    const restore = (s: typeof before): void => {
      sheet.styles = new Map(s.styles)
      const p = sheet.layout.protection
      sheet.layout = structuredClone(s.layout)
      if (p) sheet.layout.protection = p
      else delete sheet.layout.protection
    }
    let first = true
    this.commit(
      label,
      () => {
        if (first) { first = false; return }
        engine(op.kind === 'insert')
        restore(after)
      },
      () => {
        engine(op.kind === 'delete')
        if (content) this.restoreContent(id, content)
        restore(before)
      }
    )
  }

  /** Rewrites only the rows of the sheet whose content differs from `target`. */
  private restoreContent(id: number, target: (string | number | boolean | null)[][]): void {
    const cur = this.hf.getSheetSerialized(id) as (string | number | boolean | null)[][]
    const rows = Math.max(cur.length, target.length)
    const width = Math.max(0, ...target.map((r) => r.length), ...cur.map((r) => r.length))
    const differs = (r: number): boolean => {
      const a = cur[r] ?? [], b = target[r] ?? []
      for (let c = 0; c < width; c++) if ((a[c] ?? null) !== (b[c] ?? null)) return true
      return false
    }
    this.hf.batch(() => {
      for (let r = 0; r < rows; r++) {
        if (!differs(r)) continue
        let end = r
        while (end + 1 < rows && differs(end + 1)) end++
        const block = Array.from({ length: end - r + 1 }, (_, k) => Array.from({ length: width }, (_, c) => (target[r + k]?.[c] ?? null) as never))
        this.hf.setCellContents({ sheet: id, row: r, col: 0 }, block)
        r = end
      }
    })
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

  /** Group levels follow their rows/columns; a row inserted inside a group joins it, as in Excel. */
  private shiftOutline(sheet: SheetState, axis: 'row' | 'col', at: number, delta: number): void {
    const l = sheet.layout
    const levels = axis === 'row' ? l.rowLevels : l.colLevels
    if (!levels) return
    const shifted = this.shiftKeys(levels, at, delta)
    if (delta > 0 && (levels[at - 1] ?? 0) > 0 && (levels[at] ?? 0) > 0) {
      const lv = Math.min(levels[at - 1], levels[at])
      for (let i = at; i < at + delta; i++) shifted[i] = lv
    }
    const collapsed = axis === 'row' ? l.collapsedRows : l.collapsedCols
    if (axis === 'row') { l.rowLevels = shifted; if (collapsed) l.collapsedRows = this.shiftKeys(collapsed, at, delta) }
    else { l.colLevels = shifted; if (collapsed) l.collapsedCols = this.shiftKeys(collapsed, at, delta) }
    if (!Object.keys(shifted).length) { if (axis === 'row') { delete l.rowLevels; delete l.collapsedRows } else { delete l.colLevels; delete l.collapsedCols } }
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
    sheet.layout.merges = sheet.layout.merges
      .map((m) => shiftRange(m, axis, at, delta))
      .filter((m): m is Range => m !== null && (m.r2 > m.r1 || m.c2 > m.c1)) // a merge shrunk to one cell is gone
    // Notes, comments, links, conditional formatting, validation, filter and charts follow their cells.
    shiftLayoutRefs(sheet.layout, axis, at, delta)
  }

  /** Excel only deletes rows/columns on a protected sheet when every cell in them is unlocked. */
  private canDeleteLine(sheet: SheetState, axis: 'r' | 'c', at: number, count: number): boolean {
    if (!this.allows(sheet, axis === 'r' ? 'deleteRows' : 'deleteColumns')) return false
    const used = this.usedRange(sheet)
    if (!used) return true
    for (let i = at; i < at + count; i++)
      for (let j = 0; j <= (axis === 'r' ? used.c2 : used.r2); j++)
        if (this.cellLocked(sheet, axis === 'r' ? i : j, axis === 'r' ? j : i)) return this.deny(WorkbookStore.PROTECTED_MSG)
    return true
  }

  insertRows(sheet: SheetState, at: number, count: number): void {
    if (sheet.layout.protection && !this.allows(sheet, 'insertRows')) return
    this.structural(sheet, 'Insert Rows', { kind: 'insert', dim: 'r', at, count }, () => {
      this.shiftStyles(sheet, 'r', at, count)
      this.shiftMerges(sheet, 'r', at, count)
      sheet.layout.rowHeights = this.shiftKeys(sheet.layout.rowHeights, at, count)
      sheet.layout.hiddenRows = this.shiftKeys(sheet.layout.hiddenRows, at, count)
      this.shiftOutline(sheet, 'row', at, count)
    })
    this.broadcastOp({ type: 'insertRows', sheetName: sheet.name, at, count })
  }
  deleteRows(sheet: SheetState, at: number, count: number): void {
    if (sheet.layout.protection && !this.canDeleteLine(sheet, 'r', at, count)) return
    this.structural(sheet, 'Delete Rows', { kind: 'delete', dim: 'r', at, count }, () => {
      this.shiftStyles(sheet, 'r', at, -count)
      this.shiftMerges(sheet, 'r', at, -count)
      sheet.layout.rowHeights = this.shiftKeys(sheet.layout.rowHeights, at, -count)
      sheet.layout.hiddenRows = this.shiftKeys(sheet.layout.hiddenRows, at, -count)
      this.shiftOutline(sheet, 'row', at, -count)
    })
    this.broadcastOp({ type: 'deleteRows', sheetName: sheet.name, at, count })
  }
  insertCols(sheet: SheetState, at: number, count: number): void {
    if (sheet.layout.protection && !this.allows(sheet, 'insertColumns')) return
    this.structural(sheet, 'Insert Columns', { kind: 'insert', dim: 'c', at, count }, () => {
      this.shiftStyles(sheet, 'c', at, count)
      this.shiftMerges(sheet, 'c', at, count)
      sheet.layout.colWidths = this.shiftKeys(sheet.layout.colWidths, at, count)
      sheet.layout.hiddenCols = this.shiftKeys(sheet.layout.hiddenCols, at, count)
      this.shiftOutline(sheet, 'col', at, count)
    })
    this.broadcastOp({ type: 'insertCols', sheetName: sheet.name, at, count })
  }
  deleteCols(sheet: SheetState, at: number, count: number): void {
    if (sheet.layout.protection && !this.canDeleteLine(sheet, 'c', at, count)) return
    this.structural(sheet, 'Delete Columns', { kind: 'delete', dim: 'c', at, count }, () => {
      this.shiftStyles(sheet, 'c', at, -count)
      this.shiftMerges(sheet, 'c', at, -count)
      sheet.layout.colWidths = this.shiftKeys(sheet.layout.colWidths, at, -count)
      sheet.layout.hiddenCols = this.shiftKeys(sheet.layout.hiddenCols, at, -count)
      this.shiftOutline(sheet, 'col', at, -count)
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
    if (!this.canChangeStructure()) return
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
    if (!this.canChangeStructure()) return
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
    if (!this.canChangeStructure()) return null
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
      () => { this.hf.renameSheet(this.hf.getSheetId(old)!, trimmed); s.name = trimmed; this.renameInBuiltinNames(old, trimmed) },
      () => { this.hf.renameSheet(this.hf.getSheetId(trimmed)!, old); s.name = old; this.renameInBuiltinNames(trimmed, old) }
    )
    this.broadcastOp({ type: 'renameSheet', index, name: trimmed })
    return null
  }

  private renameInBuiltinNames(from: string, to: string): void {
    const q = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`)
    const esc = from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/'/g, "''")
    const re = new RegExp(`(^|[^A-Za-z0-9_.'])(?:'${esc}'|${esc})!`, 'gi')
    for (const n of this.names) if (/^_xlnm\./i.test(n.name)) n.ref = n.ref.replace(re, (_m, pre) => pre + q(to) + '!')
  }

  moveSheet(from: number, to: number): void {
    if (!this.canChangeStructure()) return
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

  /** Home > Format > Hide & Unhide > Hide Sheet. Excel keeps at least one sheet visible. */
  hideSheet(index: number): string | null {
    if (!this.canChangeStructure()) return null
    const s = this.sheets[index]
    if (!s || s.layout.sheetHidden) return null
    if (this.sheets.filter((o) => !o.layout.sheetHidden).length <= 1) return 'A workbook must contain at least one visible worksheet.'
    const prevActive = this.activeIndex
    this.commit(
      'Hide Sheet',
      () => {
        s.layout = { ...s.layout, sheetHidden: 'hidden' }
        if (this.activeIndex === index) this.activeIndex = this.nearestVisible(index)
      },
      () => { s.layout = { ...s.layout, sheetHidden: undefined }; this.activeIndex = prevActive }
    )
    return null
  }

  unhideSheet(index: number): void {
    if (!this.canChangeStructure()) return
    const s = this.sheets[index]
    if (!s || !s.layout.sheetHidden) return
    const prevActive = this.activeIndex
    const was = s.layout.sheetHidden
    this.commit(
      'Unhide Sheet',
      () => { s.layout = { ...s.layout, sheetHidden: undefined }; this.activeIndex = index },
      () => { s.layout = { ...s.layout, sheetHidden: was }; this.activeIndex = prevActive }
    )
  }

  /** The visible sheet closest to index (right first, like Excel after hiding the active sheet). */
  nearestVisible(index: number): number {
    for (let i = index + 1; i < this.sheets.length; i++) if (!this.sheets[i].layout.sheetHidden) return i
    for (let i = index - 1; i >= 0; i--) if (!this.sheets[i].layout.sheetHidden) return i
    return index
  }

  /** Move or Copy dialog: duplicates sheet `index` (cells, formulas, styles, layout) at position `to`. */
  copySheet(index: number, to: number): void {
    if (!this.canChangeStructure()) return
    const src = this.sheets[index]
    let base = src.name.replace(/ \(\d+\)$/, '')
    let n = 2
    const taken = new Set(this.sheets.map((o) => o.name.toLowerCase()))
    let name = `${base} (${n})`
    while (taken.has(name.toLowerCase())) name = `${base} (${++n})`
    if (name.length > 31) { base = base.slice(0, 31 - ` (${n})`.length); name = `${base} (${n})` }
    const content = this.hf.getSheetSerialized(this.hfId(src))
    const prevActive = this.activeIndex
    let dst: SheetState | undefined
    this.commit(
      'Copy Sheet',
      () => {
        dst = this.addSheetRaw(name, to)
        this.hf.setSheetContent(this.hfId(dst), content)
        src.styles.forEach((st, k) => dst!.styles.set(k, { ...st }))
        dst.layout = { ...JSON.parse(JSON.stringify(src.layout)), sheetHidden: undefined }
        this.activeIndex = to
      },
      () => {
        this.hf.removeSheet(this.hfId(dst!))
        this.sheets.splice(this.sheets.indexOf(dst!), 1)
        this.activeIndex = prevActive
      }
    )
  }

  setActiveSheet(index: number): void {
    if (index < 0 || index >= this.sheets.length || index === this.activeIndex) return
    if (this.sheets[index].layout.sheetHidden) return
    this.activeIndex = index
    this.clipboard = null
    this.notify()
  }

  /** F9 / Calculate Now: recomputes volatile functions (NOW, TODAY, RAND, OFFSET...) and, in manual mode, everything. */
  recalculate(): void {
    try { this.hf.rebuildAndRecalculate() } catch { /* engine keeps its last values */ }
    if (this.frozen) this.freezeValues()
    this.calcPending = false
    this.notify()
  }

  /** Shift+F9 / Calculate Sheet: refreshes only the active sheet's formula results in manual mode. */
  recalculateSheet(): void {
    if (!this.frozen) return this.recalculate()
    try { this.hf.rebuildAndRecalculate() } catch { /* */ }
    this.freezeValues(this.hfId(this.active))
    this.notify()
  }

  /** Formulas > Calculation Options. */
  setCalcMode(mode: CalcMode): void {
    if (mode === this.calcMode) return
    this.calcMode = mode
    this.dirty = true
    if (mode === 'manual') this.freezeValues()
    else { this.frozen = null; this.calcPending = false }
    this.notify()
  }

  /** Snapshots every formula cell's current result (all sheets, or one engine sheet id). */
  private freezeValues(only?: number): void {
    if (!this.frozen || only === undefined) this.frozen = new Map()
    for (const s of this.sheets) {
      const id = this.hfId(s)
      if (only !== undefined && id !== only) continue
      const ser = this.hf.getSheetSerialized(id)
      const vals = this.hf.getSheetValues(id)
      ser.forEach((row, r) => row.forEach((f, c) => {
        if (typeof f !== 'string' || !f.startsWith('=')) return
        const v = vals[r]?.[c]
        this.frozen!.set(`${id}:${r},${c}`, { f, v: (v instanceof DetailedCellError ? { error: (v.value as never) ?? '#ERROR!' } : v) as CellValue })
      }))
    }
  }

  // ---------- selection (not undoable) ----------
  setSelection(sel: Selection, sheet: SheetState = this.active): void {
    const p = sheet.layout.protection
    if (p && (!p.allow.selectLockedCells || !p.allow.selectUnlockedCells)) {
      const locked = this.cellLocked(sheet, sel.active.r, sel.active.c)
      if (locked ? !p.allow.selectLockedCells : !p.allow.selectUnlockedCells) return // Excel refuses to move onto those cells
    }
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
        if (cell.raw) (rows[r] ??= [])[c] = this.toEngine(s, r, c, cell.raw)
        if (cell.style && Object.keys(cell.style).length) s.styles.set(k, cell.style)
      }
      if (sd.packed) {
        eachPacked(sd.packed, (r, c, raw, style) => {
          if (raw) (rows[r] ??= [])[c] = this.toEngine(s, r, c, raw)
          if (style) s.styles.set(cellKey(r, c), style)
        })
      }
      // HyperFormula wants no holes in the outer array.
      for (let r = 0; r < rows.length; r++) rows[r] ??= []
      contents[sd.name] = rows
      this.sheets.push(s)
    })
    this.hf.destroy()
    this.hf = instrumentHf(data.sheets.length ? HyperFormula.buildFromSheets(contents, HF_CONFIG, BOOLEAN_NAMES) : HyperFormula.buildEmpty(HF_CONFIG, BOOLEAN_NAMES))
    // Local names are tracked by sheet identity so moving/deleting sheets can't reassign them.
    this.names = (data.names ?? []).map((n) => ({ ...n, sheetUid: n.sheet !== undefined ? this.sheets[n.sheet]?.uid : undefined }))
    this.registerNames()
    if (!this.sheets.length) this.addSheetRaw('Sheet1', 0)
    this.activeIndex = Math.min(data.activeSheet ?? 0, this.sheets.length - 1)
    if (this.sheets[this.activeIndex]?.layout.sheetHidden) {
      const v = this.sheets.findIndex((sh) => !sh.layout.sheetHidden)
      if (v >= 0) this.activeIndex = v
      else this.sheets[0].layout.sheetHidden = undefined
    }
    this.bookProtection = data.protection
    this.calcMode = data.calcMode ?? 'auto'
    this.defaultFont = data.defaultFont
    normalFont.name = data.defaultFont?.name ?? DEFAULT_FONT
    normalFont.size = data.defaultFont?.size ?? DEFAULT_FONT_SIZE
    this.frozen = null
    this.calcPending = false
    if (this.calcMode === 'manual') this.freezeValues()
    this.fileName = fileName
    this.filePath = filePath
    this.dirty = false
    this.clipboard = null
    this.undoStack = []
    this.redoStack = []
    this.notify()
  }

  /** Formulas > Define Name / Name Manager. Returns Excel's error text, or null on success. */
  defineName(name: string, refersTo: string, sheetUid?: number, replacing?: { name: string; sheetUid?: number }): string | null {
    const nm = name.trim()
    if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(nm) || nm.length > 255 || /^[A-Za-z]{1,3}[0-9]+$/.test(nm) || /^[Rr]([0-9]*)?[Cc]?([0-9]*)?$/.test(nm) || /^[CcRr]$/.test(nm) || /^(TRUE|FALSE)$/i.test(nm))
      return 'The name that you entered is not valid.\n\nReasons for this can include:\n• The name does not begin with a letter or underscore\n• The name contains a space or other invalid characters\n• The name conflicts with an Excel built-in name or the name of another object in the workbook'
    const clash = this.names.find((n) => n.name.toLowerCase() === nm.toLowerCase() && n.sheetUid === sheetUid && !(replacing && n.name === replacing.name && n.sheetUid === replacing.sheetUid))
    if (clash) return 'The name that you entered already exists. Enter a unique name.'
    const ref = refersTo.trim().replace(/^=/, '')
    if (!ref) return 'The formula contains an error.'
    const scopeOf = (uid?: number) => { const sh = uid !== undefined ? this.sheetByUid(uid) : undefined; return sh ? this.hfId(sh) : undefined }
    const old = replacing ? this.names.find((n) => n.name === replacing.name && n.sheetUid === replacing.sheetUid) : undefined
    const oldRef = old ? (() => { try { return String(this.hf.getNamedExpressionFormula(old.name, scopeOf(old.sheetUid)) ?? '=' + old.ref).replace(/^=/, '') } catch { return old.ref } })() : ''
    // Check the engine accepts it before touching anything.
    try {
      if (old) this.hf.removeNamedExpression(old.name, scopeOf(old.sheetUid))
      this.hf.addNamedExpression(nm, '=' + ref, scopeOf(sheetUid))
    } catch {
      try { if (old) this.hf.addNamedExpression(old.name, '=' + oldRef, scopeOf(old.sheetUid)) } catch { /* ignore */ }
      return 'The formula contains an error.'
    }
    try { this.hf.removeNamedExpression(nm, scopeOf(sheetUid)) } catch { /* ignore */ }
    if (old) try { this.hf.addNamedExpression(old.name, '=' + oldRef, scopeOf(old.sheetUid)) } catch { /* ignore */ }
    const entry = { name: nm, ref, sheetUid }
    const oldEntry = old ? { ...old, ref: oldRef } : undefined
    this.commit(
      old ? 'Edit Name' : 'Define Name',
      () => {
        if (oldEntry) { this.names = this.names.filter((n) => !(n.name === oldEntry.name && n.sheetUid === oldEntry.sheetUid)); try { this.hf.removeNamedExpression(oldEntry.name, scopeOf(oldEntry.sheetUid)) } catch { /* */ } }
        this.names = [...this.names, entry]
        try { this.hf.addNamedExpression(entry.name, '=' + entry.ref, scopeOf(entry.sheetUid)) } catch { /* */ }
      },
      () => {
        this.names = this.names.filter((n) => !(n.name === entry.name && n.sheetUid === entry.sheetUid))
        try { this.hf.removeNamedExpression(entry.name, scopeOf(entry.sheetUid)) } catch { /* */ }
        if (oldEntry) { this.names = [...this.names, oldEntry]; try { this.hf.addNamedExpression(oldEntry.name, '=' + oldEntry.ref, scopeOf(oldEntry.sheetUid)) } catch { /* */ } }
      }
    )
    return null
  }

  /** Page Layout > Print Area > Set / Clear (Excel's _xlnm.Print_Area name for the sheet). */
  setPrintArea(sheet: SheetState, ref: string | null): void {
    const before = this.names
    const rest = before.filter((n) => !(/^_xlnm\.Print_Area$/i.test(n.name) && n.sheetUid === sheet.uid))
    const after = ref ? [...rest, { name: '_xlnm.Print_Area', ref, sheetUid: sheet.uid }] : rest
    this.commit(ref ? 'Set Print Area' : 'Clear Print Area', () => { this.names = after }, () => { this.names = before })
  }

  /** The sheet's print title rows / columns (Excel's _xlnm.Print_Titles), as 0-based inclusive spans. */
  printTitles(sheet: SheetState): { rows?: [number, number]; cols?: [number, number] } {
    const n = this.names.find((x) => /^_xlnm\.Print_Titles$/i.test(x.name) && x.sheetUid === sheet.uid)
    const out: { rows?: [number, number]; cols?: [number, number] } = {}
    if (!n) return out
    const col = (t: string): number => [...t].reduce((v, ch) => v * 26 + ch.charCodeAt(0) - 64, 0) - 1
    for (const part of n.ref.split(',')) {
      const body = part.includes('!') ? part.slice(part.lastIndexOf('!') + 1) : part
      let m = /^\$?(\d+):\$?(\d+)$/.exec(body)
      if (m) { out.rows = [+m[1] - 1, +m[2] - 1]; continue }
      m = /^\$?([A-Za-z]+):\$?([A-Za-z]+)$/.exec(body)
      if (m) out.cols = [col(m[1].toUpperCase()), col(m[2].toUpperCase())]
    }
    return out
  }

  /** Page Layout > Print Titles: rows and/or columns repeated on every printed page (null clears). */
  setPrintTitles(sheet: SheetState, titles: { rows?: [number, number] | null; cols?: [number, number] | null }): void {
    const colName = (c: number): string => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26) } return s }
    const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet.name) ? sheet.name : `'${sheet.name.replace(/'/g, "''")}'`
    const parts: string[] = []
    if (titles.cols) parts.push(`${q}!$${colName(titles.cols[0])}:$${colName(titles.cols[1])}`)
    if (titles.rows) parts.push(`${q}!$${titles.rows[0] + 1}:$${titles.rows[1] + 1}`)
    const before = this.names
    const rest = before.filter((n) => !(/^_xlnm\.Print_Titles$/i.test(n.name) && n.sheetUid === sheet.uid))
    const after = parts.length ? [...rest, { name: '_xlnm.Print_Titles', ref: parts.join(','), sheetUid: sheet.uid }] : rest
    this.commit('Set Print Titles', () => { this.names = after }, () => { this.names = before })
  }

  deleteName(name: string, sheetUid?: number): void {
    const scopeOf = (uid?: number) => { const sh = uid !== undefined ? this.sheetByUid(uid) : undefined; return sh ? this.hfId(sh) : undefined }
    const old = this.names.find((n) => n.name === name && n.sheetUid === sheetUid)
    if (!old) return
    let ref = old.ref
    try { const f = this.hf.getNamedExpressionFormula(name, scopeOf(sheetUid)); if (typeof f === 'string') ref = f.replace(/^=/, '') } catch { /* */ }
    const entry = { ...old, ref }
    this.commit(
      'Delete Name',
      () => { this.names = this.names.filter((n) => n !== entry && !(n.name === entry.name && n.sheetUid === entry.sheetUid)); try { this.hf.removeNamedExpression(entry.name, scopeOf(entry.sheetUid)) } catch { /* */ } },
      () => { this.names = [...this.names, entry]; try { this.hf.addNamedExpression(entry.name, '=' + entry.ref, scopeOf(entry.sheetUid)) } catch { /* */ } }
    )
  }

  /** Current "Refers to" text of a name (tracks row/column edits and sheet renames). */
  nameRef(name: string, sheetUid?: number): string {
    const n = this.names.find((x) => x.name === name && x.sheetUid === sheetUid)
    if (!n) return ''
    const sh = sheetUid !== undefined ? this.sheetByUid(sheetUid) : undefined
    try { const f = this.hf.getNamedExpressionFormula(name, sh ? this.hfId(sh) : undefined); if (typeof f === 'string') return f } catch { /* */ }
    return '=' + n.ref
  }

  /** Makes ordinary defined names usable in formulas. Built-in _xlnm.* names (print areas, filter
   *  ranges) are only carried through to the saved file. A name the engine rejects stays saved. */
  private registerNames(): void {
    if (!this.names.length) return
    this.hf.batch(() => {
      for (const n of this.names) {
        if (/^_xlnm\./i.test(n.name)) continue
        const sh = n.sheetUid !== undefined ? this.sheetByUid(n.sheetUid) : undefined
        const scope = sh ? this.hfId(sh) : undefined
        try { this.hf.addNamedExpression(n.name, '=' + n.ref, scope) } catch { /* invalid for the engine */ }
      }
    })
  }

  /** Current defined names for saving: engine-tracked names reflect sheet renames and row/column
   *  edits; names local to a deleted sheet are dropped (as in Excel). */
  private exportNames(): DefinedNameData[] | undefined {
    const out: DefinedNameData[] = []
    for (const n of this.names) {
      let sheetIdx: number | undefined
      let scope: number | undefined
      if (n.sheetUid !== undefined) {
        const sh = this.sheetByUid(n.sheetUid)
        if (!sh) continue
        sheetIdx = this.sheets.indexOf(sh)
        scope = this.hfId(sh)
      }
      let ref = n.ref
      if (!/^_xlnm\./i.test(n.name)) {
        try {
          const f = this.hf.getNamedExpressionFormula(n.name, scope)
          if (typeof f === 'string') ref = f.replace(/^=/, '')
        } catch { /* keep stored ref */ }
      }
      const e: DefinedNameData = { name: n.name, ref }
      if (sheetIdx !== undefined) e.sheet = sheetIdx
      if (n.hidden) e.hidden = true
      out.push(e)
    }
    return out.length ? out : undefined
  }

  toData(): WorkbookData {
    return {
      activeSheet: this.activeIndex,
      names: this.exportNames(),
      protection: this.bookProtection,
      calcMode: this.calcMode === 'auto' ? undefined : this.calcMode,
      defaultFont: this.defaultFont,
      sheets: this.sheets.map((s) => {
        const cells: WorkbookData['sheets'][number]['cells'] = {}
        const id = this.hfId(s)
        const content = this.hf.getSheetSerialized(id)
        content.forEach((row, r) =>
          row.forEach((v, c) => {
            if (v === null || v === undefined || v === '') return
            // A spilled result (e.g. FILTER's rows below the formula) is not data: saving it as a value
            // would freeze it (losing leading zeros) and block the spill when the file is reopened.
            if (!(typeof v === 'string' && v.startsWith('=')) && this.hf.isCellPartOfArray({ sheet: id, row: r, col: c })) return
            let raw = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)
            // Typed "1/2/2024", "$1,200", "25%" are numbers to the engine: save the number.
            if (typeof v === 'string' && !v.startsWith('=') && !v.startsWith("'") && !/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(v)) {
              const val = this.hf.getCellValue({ sheet: id, row: r, col: c })
              if (typeof val === 'number') raw = String(val)
              else if (typeof val === 'boolean') raw = val ? 'TRUE' : 'FALSE'
            }
            const le = this.letSrc.get(`${s.uid}:${r},${c}`)
            if (le && le.expanded === raw) raw = le.orig
            const cell: { raw: string; value?: CellValue; spill?: string } = { raw }
            if (raw.startsWith('=')) {
              cell.value = this.getValue(s, r, c)
              const inArray = (rr: number, cc: number): boolean =>
                this.hf.isCellPartOfArray({ sheet: id, row: rr, col: cc }) && !this.hf.doesCellHaveFormula({ sheet: id, row: rr, col: cc })
              let r2 = r
              let c2 = c
              while (inArray(r2 + 1, c)) r2++
              while (inArray(r, c2 + 1)) c2++
              // HyperFormula reserves FILTER's maximum size; the real result ends at the last non-empty row/column.
              const empty = (rr: number, cc: number): boolean => { const x = this.hf.getCellValue({ sheet: id, row: rr, col: cc }); return x === null || x === '' }
              while (r2 > r && Array.from({ length: c2 - c + 1 }, (_, k) => c + k).every((cc) => empty(r2, cc))) r2--
              while (c2 > c && Array.from({ length: r2 - r + 1 }, (_, k) => r + k).every((rr) => empty(rr, c2))) c2--
              if (r2 > r || c2 > c) cell.spill = rangeA1({ r1: r, c1: c, r2, c2 })
            }
            cells[cellKey(r, c)] = cell
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
setRowVisibility((sheet, row) => store.rowHiddenKind(sheet, row))

/** React hook: re-renders the component whenever the store changes. Returns the store. */
export function useStore(): WorkbookStore {
  useSyncExternalStore(store.subscribe, store.getVersion)
  return store
}

// Handy for debugging in DevTools.
;(globalThis as unknown as { store: WorkbookStore }).store = store

// The workbook store: single source of truth for the app.
// Cell contents and formula evaluation live in HyperFormula; styles, layout and selection live here.
// Every mutation goes through `commit`, which records an undo entry and notifies subscribers.
import { HyperFormula, DetailedCellError } from 'hyperformula'
import { useSyncExternalStore } from 'react'
import {
  CellStyle, CellValue, Range, Selection, SheetLayout, cellKey, normRange,
  DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT
} from './types'
import { formatValue, detectInputFormat, FormattedValue } from '../format/numfmt'

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

export function defaultLayout(): SheetLayout {
  return {
    colWidths: {}, rowHeights: {}, hiddenRows: {}, hiddenCols: {}, merges: [],
    frozen: { rows: 0, cols: 0 }, showGridlines: true, showHeadings: true, zoom: 100
  }
}

function defaultSelection(): Selection {
  return { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }
}

let nextUid = 1

// US-locale input parsing, matching Excel en-US.
const HF_CONFIG = {
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
  clipboard: { sheetUid: number; range: Range; cut: boolean } | null = null
  /** Transient UI state any component may read (e.g. edit mode for the status bar). */
  mode: 'Ready' | 'Enter' | 'Edit' | 'Point' = 'Ready'

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
    this.hf = HyperFormula.buildEmpty(HF_CONFIG)
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
    try {
      fn() // not batched: fn may read cells between its commits
    } finally {
      this.group = null
    }
    if (!g.length) return
    this.undoStack.push({
      label,
      undo: () => this.batch(() => [...g].reverse().forEach((e) => e.undo())),
      redo: () => this.batch(() => g.forEach((e) => e.redo()))
    })
    this.redoStack = []
    this.notify()
  }

  undo(): void {
    const e = this.undoStack.pop()
    if (!e) return
    e.undo()
    this.redoStack.push(e)
    this.dirty = true
    this.notify()
  }
  redo(): void {
    const e = this.redoStack.pop()
    if (!e) return
    e.redo()
    this.undoStack.push(e)
    this.dirty = true
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
      this.batch(() => list.forEach((x) => {
        this.writeRaw(sheet, x.r, x.c, x.raw)
        this.writeStyle(sheet, x.r, x.c, x.style)
      }))
    this.commit(label, () => put(after), () => put(before))
  }

  setRaw(sheet: SheetState, r: number, c: number, raw: string): void {
    this.setRangeRaw(sheet, r, c, [[raw]])
  }

  /** Writes full cell records (raw + style) — used by paste, fill and sort. */
  setCells(sheet: SheetState, cells: { r: number; c: number; raw: string; style?: CellStyle }[], label: string): void {
    const before = cells.map((x) => ({ r: x.r, c: x.c, raw: this.getRaw(sheet, x.r, x.c), style: this.getStyle(sheet, x.r, x.c) }))
    const put = (list: typeof cells) =>
      this.batch(() => list.forEach((x) => {
        this.writeRaw(sheet, x.r, x.c, x.raw)
        this.writeStyle(sheet, x.r, x.c, x.style)
      }))
    this.commit(label, () => put(cells), () => put(before))
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
  }
  deleteRows(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Delete Rows', () => {
      this.hf.removeRows(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'r', at, -count)
      this.shiftMerges(sheet, 'r', at, -count)
      sheet.layout.rowHeights = this.shiftKeys(sheet.layout.rowHeights, at, -count)
      sheet.layout.hiddenRows = this.shiftKeys(sheet.layout.hiddenRows, at, -count)
    })
  }
  insertCols(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Insert Columns', () => {
      this.hf.addColumns(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'c', at, count)
      this.shiftMerges(sheet, 'c', at, count)
      sheet.layout.colWidths = this.shiftKeys(sheet.layout.colWidths, at, count)
      sheet.layout.hiddenCols = this.shiftKeys(sheet.layout.hiddenCols, at, count)
    })
  }
  deleteCols(sheet: SheetState, at: number, count: number): void {
    this.structural(sheet, 'Delete Columns', () => {
      this.hf.removeColumns(this.hfId(sheet), [at, count])
      this.shiftStyles(sheet, 'c', at, -count)
      this.shiftMerges(sheet, 'c', at, -count)
      sheet.layout.colWidths = this.shiftKeys(sheet.layout.colWidths, at, -count)
      sheet.layout.hiddenCols = this.shiftKeys(sheet.layout.hiddenCols, at, -count)
    })
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
    this.hf.destroy()
    this.hf = HyperFormula.buildEmpty(HF_CONFIG)
    this.sheets = []
    // Add all sheets first so cross-sheet formulas resolve.
    data.sheets.forEach((sd, i) => {
      const s = this.addSheetRaw(sd.name, i)
      s.layout = { ...defaultLayout(), ...sd.layout }
    })
    this.batch(() => {
      data.sheets.forEach((sd, i) => {
        const s = this.sheets[i]
        for (const [k, cell] of Object.entries(sd.cells)) {
          const [r, c] = k.split(',').map(Number)
          if (cell.raw) this.writeRaw(s, r, c, cell.raw)
          if (cell.style) this.writeStyle(s, r, c, cell.style)
        }
      })
    })
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

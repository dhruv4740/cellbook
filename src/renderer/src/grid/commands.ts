// Commands that act on the active sheet's current selection. Called from the grid (keyboard,
// right-click menus) and from the ribbon, so both stay consistent.
import { store, SheetState } from '../model/store'
import { setUi, getUi } from '../chrome/ui'
import { CellStyle, Range, normRange, colName, MAX_ROWS, MAX_COLS, parseA1 } from '../model/types'
import { shiftFormula } from '../engine/shiftFormula'
import { generateSeries } from '../engine/autofill'
import { sortRange, detectHeaders, hasAdjacentData, findCurrentRegion } from '../data/sort'
import { allTables } from '../model/tables'
import { commitEdit } from '../chrome/ui'

export type PasteKind = 'all' | 'values' | 'formulas' | 'formats' | 'transpose' | 'noBorders' | 'columnWidths' | 'link'

function sel(): Range[] {
  return store.active.selection.ranges
}
function activeRange(): Range {
  return normRange(sel()[sel().length - 1])
}

// ---------- clipboard ----------
interface ClipCell { r: number; c: number; raw: string; style?: CellStyle }

/** Intersection of a range with the sheet's used range (cells outside it are empty). */
function clampUsed(sheet: SheetState, rg: Range): Range | null {
  const n = normRange(rg)
  const used = store.usedRange(sheet)
  if (!used) return null
  const out = { r1: n.r1, c1: n.c1, r2: Math.min(n.r2, used.r2), c2: Math.min(n.c2, used.c2) }
  return out.r1 > out.r2 || out.c1 > out.c2 ? null : out
}

function readRange(rg: Range, sheet: SheetState = store.active): ClipCell[] {
  const n = clampUsed(sheet, rg)
  const out: ClipCell[] = []
  if (!n) return out
  for (let r = n.r1; r <= n.r2; r++)
    for (let c = n.c1; c <= n.c2; c++)
      out.push({ r, c, raw: store.getRaw(sheet, r, c), style: store.getStyle(sheet, r, c) })
  return out
}

function toTSV(rg: Range): string {
  const n = clampUsed(store.active, rg)
  if (!n) return ''
  const rows: string[] = []
  for (let r = n.r1; r <= n.r2; r++) {
    const cells: string[] = []
    for (let c = n.c1; c <= n.c2; c++) cells.push(store.getDisplay(store.active, r, c).text)
    rows.push(cells.join('\t'))
  }
  return rows.join('\n')
}

function toHTML(rg: Range): string {
  const n = clampUsed(store.active, rg)
  if (!n) return '<table></table>'
  let html = '<table>'
  for (let r = n.r1; r <= n.r2; r++) {
    html += '<tr>'
    for (let c = n.c1; c <= n.c2; c++) {
      const style = store.getStyle(store.active, r, c)
      const disp = store.getDisplay(store.active, r, c)
      const css: string[] = []
      if (style?.bold) css.push('font-weight:bold')
      if (style?.italic) css.push('font-style:italic')
      if (style?.color) css.push('color:' + style.color)
      if (style?.fill) css.push('background-color:' + style.fill)
      html += `<td style="${css.join(';')}">${disp.text.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</td>`
    }
    html += '</tr>'
  }
  return html + '</table>'
}

async function writeSystemClipboard(rg: Range): Promise<void> {
  try {
    const tsv = toTSV(rg)
    const html = toHTML(rg)
    if (navigator.clipboard && 'write' in navigator.clipboard && typeof ClipboardItem !== 'undefined') {
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob([tsv], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' })
        })
      ])
    } else if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(tsv)
    }
  } catch {
    // Clipboard API unavailable (no permission / non-secure context) — internal clipboard still works.
  }
}

export async function cut(): Promise<void> {
  const rg = activeRange()
  store.clipboard = { sheetUid: store.active.uid, range: rg, cut: true }
  await writeSystemClipboard(rg)
  store.notify()
}

export async function copy(): Promise<void> {
  const rg = activeRange()
  store.clipboard = { sheetUid: store.active.uid, range: rg, cut: false }
  await writeSystemClipboard(rg)
  store.notify()
}

function pasteCellsAt(dest: Range, src: ClipCell[], srcRange: Range, kind: PasteKind, srcSheet: SheetState = store.active): void {
  const sheet = store.active
  const n = normRange(srcRange)
  const transpose = kind === 'transpose'
  const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
  for (const cell of src) {
    const ro = cell.r - n.r1
    const co = cell.c - n.c1
    const dr = dest.r1 + (transpose ? co : ro)
    const dc = dest.c1 + (transpose ? ro : co)
    if (dr >= MAX_ROWS || dc >= MAX_COLS) continue
    let raw = cell.raw
    if (kind === 'formats' || kind === 'columnWidths') raw = store.getRaw(sheet, dr, dc)
    else if (kind === 'values') {
      const v = store.getValue(srcSheet, cell.r, cell.c)
      raw = v === null || v === undefined ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'object' ? (v as { error: string }).error : String(v)
    } else if (kind === 'link') {
      const prefix = srcSheet === sheet ? '' : "'" + srcSheet.name.replace(/'/g, "''") + "'!"
      raw = '=' + prefix + colName(cell.c) + (cell.r + 1)
    } else if (raw.startsWith('=')) {
      raw = shiftFormula(raw, dr - cell.r, dc - cell.c)
    }
    let style = kind === 'values' ? store.getStyle(sheet, dr, dc) : cell.style
    if (kind === 'formulas') style = store.getStyle(sheet, dr, dc)
    cells.push({ r: dr, c: dc, raw: kind === 'formats' ? store.getRaw(sheet, dr, dc) : raw, style })
  }
  store.setCells(sheet, cells, 'Paste')
}

async function readSystemClipboardText(): Promise<string | null> {
  try {
    return (await navigator.clipboard?.readText?.()) ?? null
  } catch {
    return null
  }
}

export async function paste(kind: PasteKind = 'all'): Promise<void> {
  const sheet = store.active
  const dest = activeRange()
  const clip = store.clipboard
  if (clip && clip.sheetUid !== undefined) {
    const srcSheet = store.sheetByUid(clip.sheetUid)
    if (srcSheet) {
      const src = readRange(clip.range, srcSheet)
      store.transact('Paste', () => {
        if (clip.cut) {
          // Clear the source first so an overlapping destination keeps the pasted cells.
          const n = clampUsed(srcSheet, clip.range)
          const clears: { r: number; c: number; raw: string }[] = []
          if (n) for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) clears.push({ r, c, raw: '' })
          store.setCells(srcSheet, clears, 'Paste')
          store.clipboard = null
        }
        pasteCellsAt(dest, src, clip.range, clip.cut ? 'all' : kind, srcSheet)
      })
      return
    }
  }
  // Fall back to the system clipboard (TSV) for cross-app paste.
  const text = await readSystemClipboardText()
  if (!text) return
  const rows = text.replace(/\r/g, '').split('\n')
  while (rows.length && rows[rows.length - 1] === '') rows.pop()
  const raws = rows.map((row) => row.split('\t'))
  store.setRangeRaw(sheet, dest.r1, dest.c1, raws, 'Paste')
}

// ---------- insert / delete ----------
export function insertCellsDialog(): void {
  setUi({ dialog: { kind: 'insertCells' } })
}
export function deleteCellsDialog(): void {
  setUi({ dialog: { kind: 'deleteCells' } })
}

function selectedRows(): number[] {
  const rows = new Set<number>()
  sel().forEach((rg) => { const n = normRange(rg); for (let r = n.r1; r <= n.r2; r++) rows.add(r) })
  return [...rows].sort((a, b) => a - b)
}
function selectedCols(): number[] {
  const cols = new Set<number>()
  sel().forEach((rg) => { const n = normRange(rg); for (let c = n.c1; c <= n.c2; c++) cols.add(c) })
  return [...cols].sort((a, b) => a - b)
}
function contiguous(nums: number[]): boolean {
  return nums.every((n, i) => i === 0 || n === nums[i - 1] + 1)
}

export function insertSheetRows(): void {
  const rows = selectedRows()
  if (!rows.length || !contiguous(rows)) return
  store.insertRows(store.active, rows[0], rows.length)
}
export function insertSheetColumns(): void {
  const cols = selectedCols()
  if (!cols.length || !contiguous(cols)) return
  store.insertCols(store.active, cols[0], cols.length)
}
export function deleteSheetRows(): void {
  const rows = selectedRows()
  if (!rows.length || !contiguous(rows)) return
  store.deleteRows(store.active, rows[0], rows.length)
}
export function deleteSheetColumns(): void {
  const cols = selectedCols()
  if (!cols.length || !contiguous(cols)) return
  store.deleteCols(store.active, cols[0], cols.length)
}

/** Shifts a block of cells (contents + styles) by (dr,dc), used by Insert/Delete "shift cells". */
function moveBlock(rg: Range, dr: number, dc: number): void {
  const sheet = store.active
  const n = normRange(rg)
  const src = readRange(n)
  store.transact('Move Cells', () => {
    // Clear the source first, then write the shifted cells.
    const clears: { r: number; c: number; raw: string }[] = src.map((x) => ({ r: x.r, c: x.c, raw: '' }))
    store.setCells(sheet, clears, 'Move Cells')
    const cells = src
      .map((cell) => ({
        r: cell.r + dr,
        c: cell.c + dc,
        raw: cell.raw.startsWith('=') ? shiftFormula(cell.raw, dr, dc) : cell.raw,
        style: cell.style
      }))
      .filter((c) => c.r >= 0 && c.c >= 0 && c.r < MAX_ROWS && c.c < MAX_COLS)
    store.setCells(sheet, cells, 'Move Cells')
  })
}

export function insertShift(kind: 'right' | 'down' | 'row' | 'col'): void {
  const rg = activeRange()
  if (kind === 'row') { insertSheetRows(); return }
  if (kind === 'col') { insertSheetColumns(); return }
  const width = rg.c2 - rg.c1 + 1
  const height = rg.r2 - rg.r1 + 1
  const bound = kind === 'right' ? { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: MAX_COLS - 1 - width } : { r1: rg.r1, c1: rg.c1, r2: MAX_ROWS - 1 - height, c2: rg.c2 }
  moveBlock(bound, kind === 'right' ? 0 : height, kind === 'right' ? width : 0)
}
export function deleteShift(kind: 'left' | 'up' | 'row' | 'col'): void {
  const rg = activeRange()
  if (kind === 'row') { deleteSheetRows(); return }
  if (kind === 'col') { deleteSheetColumns(); return }
  const width = rg.c2 - rg.c1 + 1
  const height = rg.r2 - rg.r1 + 1
  // Clear the deleted cells, then move everything after them left/up into the gap.
  const trailing = kind === 'left' ? { r1: rg.r1, c1: rg.c2 + 1, r2: rg.r2, c2: MAX_COLS - 1 } : { r1: rg.r2 + 1, c1: rg.c1, r2: MAX_ROWS - 1, c2: rg.c2 }
  const deleted = readRange(rg).map((x) => ({ r: x.r, c: x.c, raw: '' }))
  store.transact('Delete', () => {
    store.setCells(store.active, deleted, 'Delete')
    if (trailing.c1 < MAX_COLS && trailing.r1 < MAX_ROWS) moveBlock(trailing, kind === 'left' ? 0 : -height, kind === 'left' ? -width : 0)
  })
}

// ---------- clear / sort ----------
export function clearSelection(what: 'all' | 'contents' | 'formats'): void {
  store.clear(store.active, sel(), what)
}

export function openSortDialog(): void {
  const rg = activeRange()
  const sheet = store.active
  const n = normRange(rg)
  if (hasAdjacentData(store, sheet, n)) {
    setUi({
      dialog: {
        kind: 'sortWarning',
        props: {
          onSort: (expand: boolean) => {
            if (expand) {
              const exp = findCurrentRegion(store, sheet, n)
              store.setSelection({ ranges: [exp], active: sheet.selection.active })
            }
            setUi({ dialog: { kind: 'sort' } })
          }
        }
      }
    })
  } else {
    setUi({ dialog: { kind: 'sort' } })
  }
}

export function sortSelection(order: 'asc' | 'desc'): void {
  const rg = activeRange()
  const sheet = store.active
  const n = clampUsed(sheet, rg)
  if (!n) return

  if (hasAdjacentData(store, sheet, n)) {
    setUi({
      dialog: {
        kind: 'sortWarning',
        props: {
          onSort: (expand: boolean) => {
            const target = expand ? findCurrentRegion(store, sheet, n) : n
            if (expand) {
              store.setSelection({ ranges: [target], active: sheet.selection.active })
            }
            const activeCol = sheet.selection.active.c
            const hasHeaders = detectHeaders(store, sheet, target)
            sortRange(store, sheet, target, {
              hasHeaders,
              levels: [{ columnIndex: activeCol, sortOn: 'values', order }]
            })
          }
        }
      }
    })
    return
  }

  const activeCol = sheet.selection.active.c
  const hasHeaders = detectHeaders(store, sheet, n)
  sortRange(store, sheet, n, {
    hasHeaders,
    levels: [{ columnIndex: activeCol, sortOn: 'values', order }]
  })
}

// ---------- fill ----------
export function fill(direction: 'down' | 'right' | 'up' | 'left'): void {
  const rg = activeRange()
  const sheet = store.active
  const n = normRange(rg)
  const vertical = direction === 'down' || direction === 'up'
  if (vertical ? n.r1 === n.r2 : n.c1 === n.c2) return
  const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
  if (vertical) {
    for (let c = n.c1; c <= n.c2; c++) {
      const srcR = direction === 'down' ? n.r1 : n.r2
      const raw = store.getRaw(sheet, srcR, c)
      const style = store.getStyle(sheet, srcR, c)
      for (let r = n.r1; r <= n.r2; r++) {
        if (r === srcR) continue
        cells.push({ r, c, raw: raw.startsWith('=') ? shiftFormula(raw, r - srcR, 0) : raw, style })
      }
    }
  } else {
    for (let r = n.r1; r <= n.r2; r++) {
      const srcC = direction === 'right' ? n.c1 : n.c2
      const raw = store.getRaw(sheet, r, srcC)
      const style = store.getStyle(sheet, r, srcC)
      for (let c = n.c1; c <= n.c2; c++) {
        if (c === srcC) continue
        cells.push({ r, c, raw: raw.startsWith('=') ? shiftFormula(raw, 0, c - srcC) : raw, style })
      }
    }
  }
  store.setCells(sheet, cells, direction === 'down' || direction === 'right' ? 'Fill' : 'Fill')
}

/** Fill-handle drag: extends `seedRange` to `targetRange` (which fully contains it), one axis at a time. */
export function fillHandleDrag(seedRange: Range, targetRange: Range, copyOnly: boolean): void {
  const sheet = store.active
  const s = normRange(seedRange)
  const t = normRange(targetRange)
  const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
  const vertical = t.r2 - t.r1 + t.c2 - t.c1 > 0 && (t.r2 > s.r2 || t.r1 < s.r1)
  if (vertical) {
    const before = t.r1 < s.r1
    for (let c = s.c1; c <= s.c2; c++) {
      const seedRaws: string[] = []
      const rows = before ? range(s.r2, s.r1) : range(s.r1, s.r2)
      for (const r of rows) seedRaws.push(store.getRaw(sheet, r, c))
      const seedStyle = store.getStyle(sheet, before ? s.r1 : s.r2, c)
      const isFormula = seedRaws.every((r) => r.startsWith('='))
      const count = before ? s.r1 - t.r1 : t.r2 - s.r2
      let values: string[]
      if (isFormula && seedRaws.length) {
        const baseR = before ? s.r1 : s.r2
        values = range(1, count).map((i) => shiftFormula(seedRaws[seedRaws.length - 1], before ? -i : i, 0))
        void baseR
      } else {
        values = generateSeries(seedRaws, count, copyOnly)
        if (before) values.reverse()
      }
      for (let i = 0; i < count; i++) {
        const r = before ? t.r1 + i : s.r2 + 1 + i
        cells.push({ r, c, raw: values[i], style: seedStyle })
      }
    }
  } else {
    const before = t.c1 < s.c1
    for (let r = s.r1; r <= s.r2; r++) {
      const seedRaws: string[] = []
      const cols = before ? range(s.c2, s.c1) : range(s.c1, s.c2)
      for (const c of cols) seedRaws.push(store.getRaw(sheet, r, c))
      const seedStyle = store.getStyle(sheet, r, before ? s.c1 : s.c2)
      const isFormula = seedRaws.every((c) => c.startsWith('='))
      const count = before ? s.c1 - t.c1 : t.c2 - s.c2
      let values: string[]
      if (isFormula && seedRaws.length) {
        values = range(1, count).map((i) => shiftFormula(seedRaws[seedRaws.length - 1], 0, before ? -i : i))
      } else {
        values = generateSeries(seedRaws, count, copyOnly)
        if (before) values.reverse()
      }
      for (let i = 0; i < count; i++) {
        const c = before ? t.c1 + i : s.c2 + 1 + i
        cells.push({ r, c, raw: values[i], style: seedStyle })
      }
    }
  }
  if (cells.length) store.setCells(sheet, cells, 'Fill Series')
}
function range(from: number, to: number): number[] {
  const out: number[] = []
  if (from <= to) for (let i = from; i <= to; i++) out.push(i)
  else for (let i = from; i >= to; i--) out.push(i)
  return out
}

// ---------- autosum ----------
export function autoSum(fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN' = 'SUM'): void {
  const sheet = store.active
  const active = store.active.selection.active
  let r = active.r, c = active.c
  // Extend upward over a contiguous run of numeric cells.
  let top = r
  while (top > 0) {
    const v = store.getValue(sheet, top - 1, c)
    if (typeof v !== 'number') break
    top--
  }
  if (top === r) {
    // No data above: try leftward.
    let left = c
    while (left > 0) {
      const v = store.getValue(sheet, r, left - 1)
      if (typeof v !== 'number') break
      left--
    }
    if (left < c) {
      const ref = colName(left) + (r + 1) + ':' + colName(c - 1) + (r + 1)
      store.setRaw(sheet, r, c, `=${fn}(${ref})`)
      return
    }
  }
  const ref = colName(c) + (top + 1) + ':' + colName(c) + r
  store.setRaw(sheet, r, c, top < r ? `=${fn}(${ref})` : `=${fn}()`)
}

// ---------- selection helpers ----------
export function selectAll(): void {
  const sheet = store.active
  const used = store.usedRange(sheet) ?? { r1: 0, c1: 0, r2: 0, c2: 0 }
  store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }], active: sheet.selection.active }, sheet)
  void used
}

export function goTo(ref: string): boolean {
  // Table names (Ctrl+T / Format as Table) are visible in the Name Box and Go To, same as a
  // defined name - selecting one selects its whole range (header through last data/total row).
  const table = allTables(store).find((x) => x.table.name.toLowerCase() === ref.trim().toLowerCase())
  if (table) {
    if (table.sheet !== store.active) store.setActiveSheet(store.sheets.indexOf(table.sheet))
    store.setSelection({ ranges: [table.table.range], active: { r: table.table.range.r1, c: table.table.range.c1 } }, table.sheet)
    return true
  }
  const parts = ref.split('!')
  let target = ref
  let sheet = store.active
  if (parts.length === 2) {
    const s = store.sheets.find((s) => s.name.toLowerCase() === parts[0].replace(/^'|'$/g, '').toLowerCase())
    if (!s) return false
    sheet = s
    target = parts[1]
  }
  const rg = parseA1(target)
  if (!rg) return false
  if (sheet !== store.active) store.setActiveSheet(store.sheets.indexOf(sheet))
  store.setSelection({ ranges: [rg], active: { r: rg.r1, c: rg.c1 } }, store.sheets[store.activeIndex])
  return true
}

/** Ctrl+T / Ctrl+L / Insert > Table / Home > Format as Table: opens the Create Table dialog seeded
 *  with the selection (or, for a single cell, its current region - Excel's own default). */
export function openCreateTableDialog(styleName?: string): void {
  if (getUi().edit) commitEdit()
  const sheet = store.active
  const active = sheet.selection.active
  const rg = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
  const range = rg.r1 !== rg.r2 || rg.c1 !== rg.c2 ? rg : store.currentRegion(sheet, active.r, active.c)
  setUi({ dialog: { kind: 'createTable', props: { range, styleName } } })
}

// ---------- column/row sizing ----------
export function autoFitColumns(): void {
  const cols = selectedCols()
  const sheet = store.active
  const used = store.usedRange(sheet)
  cols.forEach((c) => {
    let max = 40
    if (used) {
      for (let r = used.r1; r <= Math.min(used.r2, used.r1 + 2000); r++) {
        const text = store.getDisplay(sheet, r, c).text
        if (text) max = Math.max(max, 7 * text.length + 12)
      }
    }
    store.setColWidth(sheet, [c], max)
  })
}
export function autoFitRows(): void {
  const rows = selectedRows()
  store.setRowHeight(store.active, rows, 20)
}
export function hideRows(hide: boolean): void {
  const rows = selectedRows()
  store.updateLayout(store.active, hide ? 'Hide' : 'Unhide', (l) => {
    rows.forEach((r) => { if (hide) l.hiddenRows[r] = true; else delete l.hiddenRows[r] })
  })
}
export function hideColumns(hide: boolean): void {
  const cols = selectedCols()
  store.updateLayout(store.active, hide ? 'Hide' : 'Unhide', (l) => {
    cols.forEach((c) => { if (hide) l.hiddenCols[c] = true; else delete l.hiddenCols[c] })
  })
}

// ---------- freeze panes ----------
export function freezePanes(kind: 'panes' | 'topRow' | 'firstColumn' | 'unfreeze'): void {
  const sheet = store.active
  const active = sheet.selection.active
  store.updateLayout(sheet, 'Freeze Panes', (l) => {
    if (kind === 'unfreeze') l.frozen = { rows: 0, cols: 0 }
    else if (kind === 'topRow') l.frozen = { rows: 1, cols: 0 }
    else if (kind === 'firstColumn') l.frozen = { rows: 0, cols: 1 }
    else l.frozen = { rows: active.r, cols: active.c }
  })
}

// ---------- notes / comments navigation ----------
function keyToRC(key: string): { r: number; c: number } {
  const [r, c] = key.split(',').map(Number)
  return { r, c }
}
function sortedKeys(map: Record<string, unknown>): { r: number; c: number }[] {
  return Object.keys(map).map(keyToRC).sort((a, b) => a.r - b.r || a.c - b.c)
}
/** Selects and opens the previous/next note or comment (by reading order) relative to the active cell. */
function stepAnnotation(map: Record<string, unknown>, dir: 1 | -1, open: (r: number, c: number) => void): void {
  const keys = sortedKeys(map)
  if (!keys.length) return
  const active = store.active.selection.active
  let idx = keys.findIndex((k) => k.r === active.r && k.c === active.c)
  if (idx === -1) idx = keys.findIndex((k) => k.r > active.r || (k.r === active.r && k.c > active.c))
  if (idx === -1) idx = dir === 1 ? 0 : keys.length - 1
  else idx = (idx + dir + keys.length) % keys.length
  const { r, c } = keys[idx]
  store.setSelection({ ranges: [{ r1: r, c1: c, r2: r, c2: c }], active: { r, c } })
  open(r, c)
}
export function previousNote(): void { stepAnnotation(store.active.layout.notes, -1, (r, c) => setUi({ noteEditor: { r, c } })) }
export function nextNote(): void { stepAnnotation(store.active.layout.notes, 1, (r, c) => setUi({ noteEditor: { r, c } })) }
export function previousComment(): void { stepAnnotation(store.active.layout.comments, -1, (r, c) => setUi({ commentPanel: { r, c } })) }
export function nextComment(): void { stepAnnotation(store.active.layout.comments, 1, (r, c) => setUi({ commentPanel: { r, c } })) }

// ---------- number format shortcuts (Ctrl+Shift+~ ! @ # $ % ^) ----------
export function applyNumFmtShortcut(kind: 'general' | 'number' | 'time' | 'date' | 'currency' | 'percent' | 'scientific'): void {
  const fmt = {
    general: 'General',
    number: '#,##0.00',
    time: 'h:mm AM/PM',
    date: 'm/d/yyyy',
    currency: '$#,##0.00;($#,##0.00)',
    percent: '0%',
    scientific: '0.00E+00'
  }[kind]
  store.applyStyle(store.active, sel(), { numFmt: fmt === 'General' ? undefined : fmt }, 'Number Format')
}

// ---------- quick borders (Ctrl+Shift+& / Ctrl+Shift+_) ----------
export function applyOutlineBorder(kind: 'box' | 'none'): void {
  const sheet = store.active
  const thin = { style: 'thin', color: '#000000' } as const
  sel().forEach((rg) => {
    const n = normRange(rg)
    store.applyStyle(sheet, [rg], (s, r, c) => {
      if (kind === 'none') { const { borders, ...rest } = s; void borders; return rest }
      const b: NonNullable<CellStyle['borders']> = { ...s.borders }
      if (r === n.r1) b.top = thin
      if (r === n.r2) b.bottom = thin
      if (c === n.c1) b.left = thin
      if (c === n.c2) b.right = thin
      return { ...s, borders: b }
    }, 'Borders')
  })
}

// ---------- hide/unhide via Ctrl+9 / Ctrl+0 / Ctrl+Shift+( / Ctrl+Shift+) ----------
export { hideRows as ctrlHideRows, hideColumns as ctrlHideColumns }

// ---------- copy value/formula from the cell above (Ctrl+Shift+" / Ctrl+') ----------
export function copyFromAbove(kind: 'value' | 'formula'): void {
  const sheet = store.active
  const active = sheet.selection.active
  if (active.r === 0) return
  if (kind === 'formula') {
    const raw = store.getRaw(sheet, active.r - 1, active.c)
    store.setRaw(sheet, active.r, active.c, raw.startsWith('=') ? shiftFormula(raw, 1, 0) : raw)
  } else {
    const v = store.getValue(sheet, active.r - 1, active.c)
    store.setRaw(sheet, active.r, active.c, v === null || v === undefined ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'object' ? (v as { error: string }).error : String(v))
  }
}

// ---------- select current region (Ctrl+Shift+*) ----------
export function selectCurrentRegion(): void {
  const sheet = store.active
  const active = sheet.selection.active
  if (store.getRaw(sheet, active.r, active.c) === '') return
  let r1 = active.r, r2 = active.r, c1 = active.c, c2 = active.c
  const has = (r: number, c: number): boolean => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS && store.getRaw(sheet, r, c) !== ''
  let grew = true
  while (grew) {
    grew = false
    if ([...Array(c2 - c1 + 1)].some((_, i) => has(r1 - 1, c1 + i))) { r1--; grew = true }
    if ([...Array(c2 - c1 + 1)].some((_, i) => has(r2 + 1, c1 + i))) { r2++; grew = true }
    if ([...Array(r2 - r1 + 1)].some((_, i) => has(r1 + i, c1 - 1))) { c1--; grew = true }
    if ([...Array(r2 - r1 + 1)].some((_, i) => has(r1 + i, c2 + 1))) { c2++; grew = true }
  }
  store.setSelection({ ranges: [{ r1, c1, r2, c2 }], active }, sheet)
}

// ---------- collapse a multi-range selection down to just the active cell (Shift+Backspace) ----------
export function collapseToActiveCell(): void {
  const sheet = store.active
  const a = sheet.selection.active
  store.setSelection({ ranges: [{ r1: a.r, c1: a.c, r2: a.r, c2: a.c }], active: a }, sheet)
}

// ---------- merge ----------
export function mergeCells(kind: 'center' | 'across' | 'merge' | 'unmerge'): void {
  const sheet = store.active
  const rg = normRange(activeRange())
  if (kind === 'unmerge') {
    store.updateLayout(sheet, 'Unmerge Cells', (l) => {
      l.merges = l.merges.filter((m) => !(m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2))
    })
    return
  }
  const hasContent = (() => {
    let n = 0
    for (let r = rg.r1; r <= rg.r2; r++)
      for (let c = rg.c1; c <= rg.c2; c++)
        if (store.getRaw(sheet, r, c) !== '') n++
    return n
  })()
  const doMerge = (ranges: Range[]): void => {
    store.transact('Merge Cells', () => {
      for (const m of ranges) {
        const n = normRange(m)
        // Keep only the upper-left value.
        const clears: { r: number; c: number; raw: string }[] = []
        for (let r = n.r1; r <= n.r2; r++)
          for (let c = n.c1; c <= n.c2; c++)
            if (r !== n.r1 || c !== n.c1) clears.push({ r, c, raw: '' })
        if (clears.length) store.setCells(sheet, clears, 'Merge Cells')
      }
      store.updateLayout(sheet, 'Merge Cells', (l) => {
        for (const m of ranges) l.merges.push(normRange(m))
      })
      if (kind === 'center') store.applyStyle(sheet, ranges, { hAlign: 'center' }, 'Merge Cells')
    })
  }
  const proceed = (): void => {
    if (kind === 'across') {
      const ranges: Range[] = []
      for (let r = rg.r1; r <= rg.r2; r++) ranges.push({ r1: r, c1: rg.c1, r2: r, c2: rg.c2 })
      doMerge(ranges)
    } else doMerge([rg])
  }
  if (hasContent > 1) {
    if (confirm('Merging cells only keeps the upper-left value and discards other values.'))
      proceed()
  } else proceed()
}

// ---------- Ctrl+[ / Ctrl+] precedents/dependents, Ctrl+Shift+{ / } all levels ----------
type HfAddr = { sheet: number; row: number; col: number }
type HfRange = { start: HfAddr; end: HfAddr }
function isHfRange(x: HfAddr | HfRange): x is HfRange {
  return 'start' in x
}
function hfToRange(x: HfAddr | HfRange): Range {
  if (isHfRange(x)) return { r1: x.start.row, c1: x.start.col, r2: x.end.row, c2: x.end.col }
  return { r1: x.row, c1: x.col, r2: x.row, c2: x.col }
}
function hfSheetId(x: HfAddr | HfRange): number {
  return isHfRange(x) ? x.start.sheet : x.sheet
}

/** Walk precedents/dependents on the active sheet only (cross-sheet refs aren't selectable here
 *  since a Range/Selection is scoped to one sheet), optionally recursing to "all levels". */
function traceCells(all: boolean, direction: 'precedents' | 'dependents'): Range[] {
  const sheet = store.active
  const sheetId = store.hfId(sheet)
  const seen = new Set<string>()
  const out: Range[] = []
  const visited = new Set<string>()
  const queue: HfAddr[] = [{ sheet: sheetId, row: sheet.selection.active.r, col: sheet.selection.active.c }]
  let first = true
  while (queue.length) {
    const addr = queue.shift()!
    const key = addr.sheet + ':' + addr.row + ':' + addr.col
    if (visited.has(key)) continue
    visited.add(key)
    let related: (HfAddr | HfRange)[]
    try {
      related = direction === 'precedents' ? store.hf.getCellPrecedents(addr) : store.hf.getCellDependents(addr)
    } catch { related = [] }
    for (const rel of related) {
      if (hfSheetId(rel) !== sheetId) continue // cross-sheet: not representable in this sheet's selection
      const rg = hfToRange(rel)
      const k = `${rg.r1},${rg.c1},${rg.r2},${rg.c2}`
      if (!seen.has(k)) { seen.add(k); out.push(rg) }
      if (all && !isHfRange(rel)) queue.push(rel)
    }
    if (!all) break
    first = false
  }
  void first
  return out
}

export function selectPrecedents(all: boolean): void {
  const ranges = traceCells(all, 'precedents')
  if (!ranges.length) return
  const a = ranges[0]
  store.setSelection({ ranges, active: { r: a.r1, c: a.c1 } }, store.active)
}
export function selectDependents(all: boolean): void {
  const ranges = traceCells(all, 'dependents')
  if (!ranges.length) return
  const a = ranges[0]
  store.setSelection({ ranges, active: { r: a.r1, c: a.c1 } }, store.active)
}

// ---------- Ctrl+/: select the current array (CSE) formula's cells ----------
export function selectCurrentArray(): void {
  const sheet = store.active
  const a = sheet.selection.active
  const raw = store.getRaw(sheet, a.r, a.c)
  if (!raw.startsWith('=')) return
  // Grow a block of cells sharing the exact same formula text (this app has no distinct CSE
  // marker, so an identical-formula contiguous block is the closest equivalent).
  let r1 = a.r, r2 = a.r, c1 = a.c, c2 = a.c
  const same = (r: number, c: number): boolean => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS && store.getRaw(sheet, r, c) === raw
  let grew = true
  while (grew) {
    grew = false
    if ([...Array(c2 - c1 + 1)].some((_, i) => same(r1 - 1, c1 + i))) { r1--; grew = true }
    if ([...Array(c2 - c1 + 1)].some((_, i) => same(r2 + 1, c1 + i))) { r2++; grew = true }
    if ([...Array(r2 - r1 + 1)].some((_, i) => same(r1 + i, c1 - 1))) { c1--; grew = true }
    if ([...Array(r2 - r1 + 1)].some((_, i) => same(r1 + i, c2 + 1))) { c2++; grew = true }
  }
  store.setSelection({ ranges: [{ r1, c1, r2, c2 }], active: a }, sheet)
}

// ---------- Alt+;: select only the visible cells in the current selection (filter-aware) ----------
export function selectVisibleCells(): void {
  const sheet = store.active
  const hiddenRows = sheet.layout.hiddenRows ?? {}
  const hiddenCols = sheet.layout.hiddenCols ?? {}
  const filteredOut = new Set(sheet.layout.autoFilter?.filteredOutRows ?? [])
  const out: Range[] = []
  for (const raw of sel()) {
    const n = normRange(raw)
    let blockStart: number | null = null
    for (let r = n.r1; r <= n.r2 + 1; r++) {
      const visible = r <= n.r2 && !hiddenRows[r] && !filteredOut.has(r)
      if (visible && blockStart === null) blockStart = r
      else if (!visible && blockStart !== null) { out.push({ r1: blockStart, c1: n.c1, r2: r - 1, c2: n.c2 }); blockStart = null }
    }
  }
  const finalRanges = out.filter((rg) => {
    for (let c = rg.c1; c <= rg.c2; c++) if (!hiddenCols[c]) return true
    return false
  })
  if (finalRanges.length) store.setSelection({ ranges: finalRanges, active: store.active.selection.active }, sheet)
}

// ---------- Alt+Down: pick-from-list ----------
/** Unique text values from the contiguous data block above/below the active cell, most-recent
 *  first, for the Alt+Down "pick from list" autocomplete. */
export function pickListValues(): string[] {
  const sheet = store.active
  const a = sheet.selection.active
  const seen = new Set<string>()
  const out: string[] = []
  for (let r = a.r - 1; r >= 0; r--) {
    const raw = store.getRaw(sheet, r, a.c)
    if (raw === '') { if (r < a.r - 1 && out.length) break; continue }
    if (raw.startsWith('=')) continue
    const disp = store.getDisplay(sheet, r, a.c).text
    if (!seen.has(disp)) { seen.add(disp); out.push(disp) }
  }
  for (let r = a.r + 1; r < MAX_ROWS; r++) {
    const raw = store.getRaw(sheet, r, a.c)
    if (raw === '') break
    if (raw.startsWith('=')) continue
    const disp = store.getDisplay(sheet, r, a.c).text
    if (!seen.has(disp)) { seen.add(disp); out.push(disp) }
  }
  return out
}

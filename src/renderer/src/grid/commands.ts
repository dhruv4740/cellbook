import { moveRefs } from '../engine/moveRefs'
// Commands that act on the active sheet's current selection. Called from the grid (keyboard,
// right-click menus) and from the ribbon, so both stay consistent.
import { store, SheetState } from '../model/store'
import { setUi, getUi } from '../chrome/ui'
import { CellStyle, Range, normRange, colName, MAX_ROWS, MAX_COLS, parseA1 } from '../model/types'
import { shiftFormula } from '../engine/shiftFormula'
import { generateSeries } from '../engine/autofill'
import { adjustDecimals } from '../format/decimals'
import { seriesValues, SeriesType, DateUnit } from '../engine/series'
import { flashFill } from '../data/flashFill'
import { sortRange, detectHeaders, hasAdjacentData, findCurrentRegion } from '../data/sort'

export type PasteKind = 'all' | 'values' | 'formulas' | 'formats' | 'transpose' | 'noBorders' | 'columnWidths' | 'link'

function sel(): Range[] {
  return store.active.selection.ranges
}
function activeRange(): Range {
  return normRange(sel()[sel().length - 1])
}

// ---------- clipboard ----------
interface ClipCell { r: number; c: number; raw: string; style?: CellStyle; pr?: number }

/** Rows a copy skips: with a filter on, Excel copies only the visible rows. */
function skippedRows(sheet: SheetState, rg: Range): Set<number> {
  const out = new Set<number>()
  if (!sheet.layout.autoFilter?.filteredOutRows?.length) return out
  const n = normRange(rg)
  for (const k of Object.keys(sheet.layout.hiddenRows ?? {})) { const r = +k; if (r >= n.r1 && r <= n.r2) out.add(r) }
  return out
}

/** Intersection of a range with the sheet's used range (cells outside it are empty). */
function clampUsed(sheet: SheetState, rg: Range): Range | null {
  const n = normRange(rg)
  const used = store.usedRange(sheet)
  if (!used) return null
  const out = { r1: n.r1, c1: n.c1, r2: Math.min(n.r2, used.r2), c2: Math.min(n.c2, used.c2) }
  return out.r1 > out.r2 || out.c1 > out.c2 ? null : out
}

function readRange(rg: Range, sheet: SheetState = store.active, skip?: Set<number>): ClipCell[] {
  const n = clampUsed(sheet, rg)
  const out: ClipCell[] = []
  if (!n) return out
  let pr = n.r1
  for (let r = n.r1; r <= n.r2; r++) {
    if (skip?.has(r)) continue
    for (let c = n.c1; c <= n.c2; c++)
      out.push({ r, c, raw: store.getRaw(sheet, r, c), style: store.getStyle(sheet, r, c), pr })
    pr++
  }
  return out
}

function toTSV(rg: Range): string {
  const n = clampUsed(store.active, rg)
  if (!n) return ''
  const rows: string[] = []
  const skip = skippedRows(store.active, n)
  for (let r = n.r1; r <= n.r2; r++) {
    if (skip.has(r)) continue
    const cells: string[] = []
    for (let c = n.c1; c <= n.c2; c++) {
      const t = store.getDisplay(store.active, r, c).text
      // Like Excel: a cell holding a tab, line break or quote is wrapped in quotes so it stays one cell.
      cells.push(/[\t\n"]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t)
    }
    rows.push(cells.join('\t'))
  }
  return rows.join('\n')
}

function toHTML(rg: Range): string {
  const n = clampUsed(store.active, rg)
  if (!n) return '<table></table>'
  let html = '<table>'
  const skip = skippedRows(store.active, n)
  for (let r = n.r1; r <= n.r2; r++) {
    if (skip.has(r)) continue
    html += '<tr>'
    for (let c = n.c1; c <= n.c2; c++) {
      const style = store.getStyle(store.active, r, c)
      const disp = store.getDisplay(store.active, r, c)
      const css: string[] = []
      if (style?.bold) css.push('font-weight:bold')
      if (style?.italic) css.push('font-style:italic')
      if (style?.color) css.push('color:' + style.color)
      if (style?.fill) css.push('background-color:' + style.fill)
      // Excel's own markers: text stays text (keeps "00012"), line breaks stay inside the cell.
      // (only on text Excel would otherwise read as a number or date, so other cells keep General).
      const v = store.getValue(store.active, r, c)
      if (typeof v === 'string' && /\d/.test(v) && /^[-+\d\s.,/:%₹$()E]+$/i.test(v)) css.push('mso-number-format:"\\@"')
      const body = disp.text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br style="mso-data-placement:same-cell;"/>')
      html += `<td style='${css.join(';')}'>${body}</td>`
    }
    html += '</tr>'
  }
  return html + '</table>'
}

async function writeSystemClipboard(rg: Range): Promise<void> {
  const tsv = toTSV(rg)
  if (store.clipboard) store.clipboard.text = tsv
  try {
    const html = toHTML(rg)
    if (window.api?.clipboardWrite) {
      await window.api.clipboardWrite({ text: tsv, html })
    } else if (navigator.clipboard && 'write' in navigator.clipboard && typeof ClipboardItem !== 'undefined') {
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
    // Clipboard API unavailable (no permission / non-secure context) â€” internal clipboard still works.
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
    const ro = (cell.pr ?? cell.r) - n.r1
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

/** Cut + Paste: moves the cells unchanged and re-points every formula that referenced them. */
function cutMove(srcSheet: SheetState, range: Range, dest: Range): void {
  const destSheet = store.active
  const n = normRange(range)
  const dr = dest.r1 - n.r1, dc = dest.c1 - n.c1
  const src = readRange(n, srcSheet)
  store.transact('Paste', () => {
    const clears: { r: number; c: number; raw: string; style?: CellStyle }[] = []
    const u = clampUsed(srcSheet, n)
    if (u) for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) clears.push({ r, c, raw: '', style: undefined })
    store.setCells(srcSheet, clears, 'Paste')
    const moved = src
      .map((cell) => ({ r: cell.r + dr, c: cell.c + dc, raw: cell.raw, style: cell.style }))
      .filter((c) => c.r < MAX_ROWS && c.c < MAX_COLS)
    store.setCells(destSheet, moved, 'Paste')
    // Re-point formulas (including the moved ones) at the cells' new home.
    for (const sh of store.sheets) {
      const formulas = store.hf.getSheetFormulas(store.hfId(sh))
      const edits: { r: number; c: number; raw: string; style?: CellStyle }[] = []
      formulas.forEach((row, r) => row?.forEach((f, c) => {
        if (!f) return
        const next = moveRefs(f, sh.name, srcSheet.name, n, dr, dc, destSheet.name)
        if (next !== f) edits.push({ r, c, raw: next, style: store.getStyle(sh, r, c) })
      }))
      if (edits.length) store.setCells(sh, edits, 'Paste')
    }
  })
  store.setSelection({ ranges: [{ r1: dest.r1, c1: dest.c1, r2: dest.r1 + n.r2 - n.r1, c2: dest.c1 + n.c2 - n.c1 }], active: { r: dest.r1, c: dest.c1 } })
}

async function readSystemClipboardText(): Promise<string | null> {
  try {
    if (window.api?.clipboardReadText) return await window.api.clipboardReadText()
    return (await navigator.clipboard?.readText?.()) ?? null
  } catch {
    return null
  }
}

export async function paste(kind: PasteKind = 'all'): Promise<void> {
  const sheet = store.active
  const dest = activeRange()
  let clip = store.clipboard
  const sysText = await readSystemClipboardText()
  // Something new was copied in another app since our Copy: Excel drops its own clipboard then too.
  // (An empty system clipboard - blocked by clipboard-protection software, say - is not "something else was copied".)
  if (clip && sysText && clip.text !== undefined && norm(sysText) !== norm(clip.text)) {
    store.clipboard = clip = null
    store.notify()
  }
  if (clip && clip.sheetUid !== undefined) {
    const srcSheet = store.sheetByUid(clip.sheetUid)
    if (srcSheet) {
      const skip = clip.cut ? undefined : skippedRows(srcSheet, clip.range)
      const src = readRange(clip.range, srcSheet, skip)
      if (clip.cut) {
        cutMove(srcSheet, clip.range, dest)
        store.clipboard = null
        store.notify()
        return
      }
      const sn = normRange(clip.range)
      const used = clampUsed(srcSheet, sn)
      const skipped = skip ? [...skip].filter((r) => r <= (used?.r2 ?? sn.r2)).length : 0
      const targets = tileTargets(dest, Math.max(1, Math.min(sn.r2, used?.r2 ?? sn.r2) - sn.r1 + 1 - skipped), sn.c2 - sn.c1 + 1, kind === 'transpose')
      store.transact('Paste', () => {
        if (clip.cut) {
          // Clear the source first so an overlapping destination keeps the pasted cells.
          const n = clampUsed(srcSheet, clip.range)
          const clears: { r: number; c: number; raw: string }[] = []
          if (n) for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) clears.push({ r, c, raw: '' })
          store.setCells(srcSheet, clears, 'Paste')
          store.clipboard = null
        }
        for (const d of targets) pasteCellsAt(d, src, clip.range, clip.cut ? 'all' : kind, srcSheet)
      })
      selectPasted(dest, targets)
      store.notify()
      return
    }
  }
  // Fall back to the system clipboard (TSV) for cross-app paste.
  if (!sysText) return
  const raws = parseTsv(sysText)
  if (!raws.length) return
  // Excel's HTML says which cells are Text-formatted; the plain text alone would turn "00012" into 12.
  const textCells = excelTextCells((await window.api?.clipboardReadHTML?.().catch(() => '')) ?? '')
  if (textCells && textCells.length === raws.length) {
    raws.forEach((row, i) => row.forEach((raw, j) => {
      if (textCells[i]?.[j] && raw !== '' && !raw.startsWith("'")) row[j] = "'" + raw
    }))
  }
  const width = Math.max(...raws.map((r) => r.length))
  const targets = tileTargets(dest, raws.length, width, false)
  store.transact('Paste', () => {
    for (const d of targets) store.setRangeRaw(sheet, d.r1, d.c1, raws, 'Paste')
  })
  selectPasted(dest, targets)
}

/** Excel leaves the pasted block selected. */
function selectPasted(dest: Range, targets: Range[]): void {
  const d = normRange(dest)
  const last = targets[targets.length - 1]
  const r2 = Math.max(d.r2, last.r2), c2 = Math.max(d.c2, last.c2)
  store.setSelection({ ranges: [{ r1: d.r1, c1: d.c1, r2: Math.min(r2, MAX_ROWS - 1), c2: Math.min(c2, MAX_COLS - 1) }], active: { r: d.r1, c: d.c1 } })
}

function norm(t: string): string {
  return t.replace(/\r/g, '').replace(/\n+$/, '')
}

/** From Excel's clipboard HTML, which cells are Text-formatted (mso-number-format:"\@" inline or via a
 *  class in its <style>). Null when there is no simple table (merged cells etc.). */
export function excelTextCells(html: string): boolean[][] | null {
  if (!html || !/<td/i.test(html)) return null
  const isText = (css: string): boolean => /mso-number-format:\s*"?\\?\\@"?/i.test(css)
  const textClasses = new Set<string>()
  for (const m of html.matchAll(/\.([\w-]+)\s*\{([^}]*)\}/g)) if (isText(m[2])) textClasses.add(m[1])
  const rows: boolean[][] = []
  for (const tr of html.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const row: boolean[] = []
    for (const td of tr[1].matchAll(/<td\b([^>]*)>/gi)) {
      const attrs = td[1]
      if (/\b(?:col|row)span\s*=\s*["']?(?:[2-9]|\d\d)/i.test(attrs)) return null
      const cls = /\bclass\s*=\s*["']?([\w-]+)/i.exec(attrs)?.[1]
      const style = /\bstyle\s*=\s*(?:'([^']*)'|"([^"]*)")/i.exec(attrs)
      row.push((cls !== undefined && textClasses.has(cls)) || isText(style?.[1] ?? style?.[2] ?? ''))
    }
    rows.push(row)
  }
  return rows.length ? rows : null
}

/** Excel's clipboard text: tab/newline separated, fields holding tabs/newlines/quotes wrapped in "..." with "" escapes. */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let i = 0
  const t = text.replace(/\r\n?/g, '\n')
  while (i < t.length) {
    const ch = t[i]
    if (field === '' && ch === '"') {
      const end = findQuoteEnd(t, i + 1)
      if (end !== -1 && (end + 1 === t.length || t[end + 1] === '\t' || t[end + 1] === '\n')) {
        field = t.slice(i + 1, end).replace(/""/g, '"')
        i = end + 1
        continue
      }
    }
    if (ch === '\t') { row.push(field); field = '' }
    else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = '' }
    else field += ch
    i++
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row) }
  return rows
}
function findQuoteEnd(t: string, from: number): number {
  for (let i = from; i < t.length; i++) {
    if (t[i] === '"') {
      if (t[i + 1] === '"') { i++; continue }
      return i
    }
  }
  return -1
}

/** Excel repeats the copied block across the selection when the selection is an exact multiple of it. */
export function tileTargets(dest: Range, h: number, w: number, transpose: boolean): Range[] {
  const d = normRange(dest)
  const bh = transpose ? w : h
  const bw = transpose ? h : w
  const dh = d.r2 - d.r1 + 1
  const dw = d.c2 - d.c1 + 1
  if (dh % bh !== 0 || dw % bw !== 0 || (dh === bh && dw === bw) || dh * dw > 1_000_000) {
    return [{ r1: d.r1, c1: d.c1, r2: d.r1 + bh - 1, c2: d.c1 + bw - 1 }]
  }
  const out: Range[] = []
  for (let r = d.r1; r <= d.r2; r += bh) for (let c = d.c1; c <= d.c2; c += bw) out.push({ r1: r, c1: c, r2: r + bh - 1, c2: c + bw - 1 })
  return out
}

// ---------- insert / delete ----------
// Whole rows/columns selected: Excel inserts/deletes straight away, no dialog.
function wholeLines(): 'rows' | 'cols' | null {
  const rs = sel().map(normRange)
  if (rs.every((r) => r.c1 === 0 && r.c2 >= MAX_COLS - 1)) return 'rows'
  if (rs.every((r) => r.r1 === 0 && r.r2 >= MAX_ROWS - 1)) return 'cols'
  return null
}
export function insertCellsDialog(): void {
  const w = wholeLines()
  if (w === 'rows') return insertSheetRows()
  if (w === 'cols') return insertSheetColumns()
  setUi({ dialog: { kind: 'insertCells' } })
}
export function deleteCellsDialog(): void {
  const w = wholeLines()
  if (w === 'rows') return deleteSheetRows()
  if (w === 'cols') return deleteSheetColumns()
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
export function clearSelection(what: 'all' | 'contents' | 'formats' | 'notes' | 'hyperlinks'): void {
  const sheet = store.active
  if (what === 'notes' || what === 'hyperlinks') {
    const ranges = sel().map(normRange)
    const inSel = (key: string) => { const [r, c] = key.split(',').map(Number); return ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) }
    if (what === 'notes') {
      store.updateLayout(sheet, 'Clear Comments and Notes', (l) => {
        l.notes = Object.fromEntries(Object.entries(l.notes).filter(([k]) => !inSel(k)))
        l.comments = Object.fromEntries(Object.entries(l.comments).filter(([k]) => !inSel(k)))
      })
    } else {
      // Excel's Clear Hyperlinks removes the link but keeps the text and its formatting.
      store.updateLayout(sheet, 'Clear Hyperlinks', (l) => {
        if (l.hyperlinks) l.hyperlinks = Object.fromEntries(Object.entries(l.hyperlinks).filter(([k]) => !inSel(k)))
      })
    }
    return
  }
  store.clear(sheet, sel(), what)
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
  if (!store.allows(store.active, 'formatRows')) return
  const rows = selectedRows()
  store.updateLayout(store.active, hide ? 'Hide' : 'Unhide', (l) => {
    rows.forEach((r) => { if (hide) l.hiddenRows[r] = true; else delete l.hiddenRows[r] })
  })
}
export function hideColumns(hide: boolean): void {
  if (!store.allows(store.active, 'formatColumns')) return
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
  if (!store.allows(sheet, 'formatCells')) return
  const rg = normRange(activeRange())
  const overlaps = (m: Range) => !(m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2)
  // Merge & Center on a merged cell unmerges it (the button is a toggle in Excel).
  const already = sheet.layout.merges.some((m) => { const x = normRange(m); return x.r1 === rg.r1 && x.r2 === rg.r2 && x.c1 === rg.c1 && x.c2 === rg.c2 })
  if (kind === 'unmerge' || (kind === 'center' && already)) {
    store.updateLayout(sheet, 'Unmerge Cells', (l) => {
      l.merges = l.merges.filter((m) => !overlaps(normRange(m)))
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
        // Existing merges touching the selection are replaced, never stacked.
        l.merges = l.merges.filter((m) => !overlaps(normRange(m)))
        for (const m of ranges) { const n = normRange(m); if (n.r1 !== n.r2 || n.c1 !== n.c2) l.merges.push(n) }
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

// ---------- Home > Format as Table / Insert > Table ----------
/** Excel's built-in table style families: [header fill, band fill, border colour]. */
export const TABLE_STYLES: { name: string; header: string; band: string; line: string; headerFont: string }[] = [
  { name: 'Blue, Table Style Medium 2', header: '#4472C4', band: '#D9E1F2', line: '#8EA9DB', headerFont: '#FFFFFF' },
  { name: 'Orange, Table Style Medium 3', header: '#ED7D31', band: '#FCE4D6', line: '#F4B084', headerFont: '#FFFFFF' },
  { name: 'Gray, Table Style Medium 4', header: '#A5A5A5', band: '#EDEDED', line: '#C9C9C9', headerFont: '#FFFFFF' },
  { name: 'Gold, Table Style Medium 5', header: '#FFC000', band: '#FFF2CC', line: '#FFD966', headerFont: '#FFFFFF' },
  { name: 'Light Blue, Table Style Medium 6', header: '#5B9BD5', band: '#DDEBF7', line: '#9BC2E6', headerFont: '#FFFFFF' },
  { name: 'Green, Table Style Medium 7', header: '#70AD47', band: '#E2EFDA', line: '#A9D08E', headerFont: '#FFFFFF' },
  { name: 'Light Blue, Table Style Light 9', header: '#4472C4', band: '#FFFFFF', line: '#4472C4', headerFont: '#FFFFFF' },
  { name: 'Black, Table Style Medium 1', header: '#000000', band: '#D9D9D9', line: '#000000', headerFont: '#FFFFFF' }
]

/**
 * Formats the selection (or the current region around the active cell) as a banded table with
 * a filter on the header row, in one undo step. Cellbook stores this as ordinary formatting plus
 * an AutoFilter; it isn't saved as an Excel table object.
 */
export function formatAsTable(styleIndex = 0): void {
  const sheet = store.active
  const n = normRange(activeRange())
  const region = n.r1 !== n.r2 || n.c1 !== n.c2 ? store.clampToUsed(sheet, n) : store.currentRegion(sheet, sheet.selection.active.r, sheet.selection.active.c)
  if (region.r2 <= region.r1 && region.c1 === region.c2 && !store.getRaw(sheet, region.r1, region.c1)) {
    setUi({ dialog: { kind: 'message', props: { text: 'Select a range with data first, then choose Format as Table.' } } })
    return
  }
  const t = TABLE_STYLES[styleIndex] ?? TABLE_STYLES[0]
  const side = { style: 'thin' as const, color: t.line }
  store.transact('Format as Table', () => {
    store.applyStyle(sheet, [region], (s, r) => {
      const next: CellStyle = { ...s, borders: { top: side, bottom: side, left: side, right: side } }
      if (r === region.r1) Object.assign(next, { bold: true, fill: t.header, color: t.headerFont })
      else next.fill = (r - region.r1) % 2 === 1 ? t.band : undefined
      return next
    })
    if (sheet.layout.autoFilter) store.toggleAutoFilter(sheet)
    store.setSelection({ ranges: [region], active: { r: region.r1, c: region.c1 } })
    store.toggleAutoFilter(sheet)
  })
}

// ---------- Home > Fill > Series... ----------
export interface SeriesOptions { byRows: boolean; type: SeriesType; unit: DateUnit; step: number; stop?: number }

/** Fills each row/column of the selection from its first cell, as Excel's Series dialog. */
export function fillSeries(o: SeriesOptions): string | null {
  const sheet = store.active
  const n = normRange(activeRange())
  const cells: { r: number; c: number; raw: string; style?: CellStyle }[] = []
  const lanes = o.byRows ? [n.r1, n.r2] : [n.c1, n.c2]
  for (let lane = lanes[0]; lane <= lanes[1]; lane++) {
    const r0 = o.byRows ? lane : n.r1, c0 = o.byRows ? n.c1 : lane
    const v = store.getValue(sheet, r0, c0)
    if (typeof v !== 'number') continue
    let count = o.byRows ? n.c2 - n.c1 : n.r2 - n.r1
    if (count === 0) {
      if (o.stop === undefined) continue
      if ((o.type === 'growth' && Math.abs(o.step) <= 1) || (o.type !== 'growth' && o.step === 0)) return 'That step value would never reach the stop value.'
      count = o.byRows ? MAX_COLS - 1 - c0 : Math.min(MAX_ROWS - 1 - r0, 1_000_000)
    }
    const vals = seriesValues(v, count, o.type, o.step, o.stop, o.unit)
    const style = store.getStyle(sheet, r0, c0)
    vals.forEach((x, k) => cells.push({ r: o.byRows ? r0 : r0 + k + 1, c: o.byRows ? c0 + k + 1 : c0, raw: String(x), style }))
  }
  if (!cells.length) return 'Select a range whose first cell contains a number or date (and, for a single cell, enter a Stop value).'
  store.setCells(sheet, cells, 'Series')
  return null
}

/** View > Zoom to Selection: zooms so the selected range fills the grid area. */
export function zoomToSelection(): void {
  const sheet = store.active
  const n = normRange(activeRange())
  let w = 0, h = 0
  for (let c = n.c1; c <= Math.min(n.c2, n.c1 + 500); c++) if (!sheet.layout.hiddenCols[c]) w += store.colWidth(sheet, c)
  for (let r = n.r1; r <= Math.min(n.r2, n.r1 + 2000); r++) if (!sheet.layout.hiddenRows[r]) h += store.rowHeight(sheet, r)
  const host = document.querySelector('.grid-host') as HTMLElement | null
  const W = (host?.clientWidth ?? window.innerWidth) - 60, H = (host?.clientHeight ?? window.innerHeight - 250) - 40
  const z = Math.max(10, Math.min(400, Math.floor(100 * Math.min(W / Math.max(1, w), H / Math.max(1, h)))))
  store.updateLayout(sheet, 'Zoom', (l) => (l.zoom = z))
}

/** Home > Increase / Decrease Decimal, based on the active cell's format (as Excel). */
export function adjustDecimal(delta: 1 | -1): void {
  const sheet = store.active
  const a = sheet.selection.active
  const fmt = adjustDecimals(store.getStyle(sheet, a.r, a.c)?.numFmt, store.getDisplay(sheet, a.r, a.c).text, delta)
  store.applyStyle(sheet, sel(), { numFmt: fmt }, delta > 0 ? 'Increase Decimal' : 'Decrease Decimal')
}

/** Home > Borders presets. Edge presets (Bottom, Top, Left, Right...) apply to the outer edge of
 *  each selected range, like Excel; All Borders applies to every cell. One undo step. */
export function applyBorderPreset(kind: string): void {
  const sheet = store.active
  const thin = { style: 'thin', color: '#000000' } as const
  const thick = { style: 'medium', color: '#000000' } as const
  const dbl = { style: 'double', color: '#000000' } as const
  store.transact('Borders', () => {
    for (const rg of sel()) {
      const n = normRange(rg)
      store.applyStyle(sheet, [rg], (s, r, c) => {
        const b: NonNullable<CellStyle['borders']> = { ...s.borders }
        const top = r === n.r1, bottom = r === n.r2, left = c === n.c1, right = c === n.c2
        switch (kind) {
          case 'bottom': if (bottom) b.bottom = thin; break
          case 'top': if (top) b.top = thin; break
          case 'left': if (left) b.left = thin; break
          case 'right': if (right) b.right = thin; break
          case 'none': delete b.top; delete b.right; delete b.bottom; delete b.left; break
          case 'all': b.top = thin; b.bottom = thin; b.left = thin; b.right = thin; break
          case 'outside': if (top) b.top = thin; if (bottom) b.bottom = thin; if (left) b.left = thin; if (right) b.right = thin; break
          case 'thickOutside': if (top) b.top = thick; if (bottom) b.bottom = thick; if (left) b.left = thick; if (right) b.right = thick; break
          case 'bottomDouble': if (bottom) b.bottom = dbl; break
          case 'thickBottom': if (bottom) b.bottom = thick; break
          case 'topBottom': if (top) b.top = thin; if (bottom) b.bottom = thin; break
          case 'topThickBottom': if (top) b.top = thin; if (bottom) b.bottom = thick; break
          case 'topDoubleBottom': if (top) b.top = thin; if (bottom) b.bottom = dbl; break
        }
        return { ...s, borders: Object.keys(b).length ? b : undefined }
      }, 'Borders')
    }
  })
}

/** Data > Flash Fill / Ctrl+E: fills the active column below the typed examples by pattern. */
export function flashFillSelection(): void {
  const sheet = store.active
  const { r, c } = sheet.selection.active
  const region = (col: number): Range | null => (col >= 0 && col < MAX_COLS && store.getRaw(sheet, r, col) !== '' ? store.currentRegion(sheet, r, col) : null)
  // Sources are the columns of the adjoining data block: left of the target, else right of it.
  const src: number[] = []
  let rows: Range | null = null
  const left = region(c - 1)
  if (left) { rows = left; for (let k = left.c1; k < c; k++) src.push(k) }
  else {
    const right = region(c + 1)
    if (right) { rows = right; for (let k = c + 1; k <= right.c2; k++) src.push(k) }
  }
  if (!rows || !src.length) return
  const r1 = rows.r1, r2 = rows.r2
  const data: string[][] = [], targets: string[] = []
  for (let rr = r1; rr <= r2; rr++) {
    data.push(src.map((k) => store.getDisplay(sheet, rr, k).text))
    targets.push(store.getRaw(sheet, rr, c))
  }
  const out = flashFill(data, targets)
  if (!out) return
  const cells = out.map((v, i) => (v === null ? null : { r: r1 + i, c, raw: v })).filter((x): x is { r: number; c: number; raw: string } => x !== null)
  if (cells.length) store.setCells(sheet, cells, 'Flash Fill')
}

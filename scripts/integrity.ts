// Data-integrity check on the real despatch workbook scaled up: every operation is compared with an independent
// model, and every undo must restore the exact raw text, value, display text and style of every cell.
//   npx vite-node scripts/integrity.ts [rows]
import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { writeXlsx, readXlsx } from '../src/renderer/src/io/xlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
import { sortRange } from '../src/renderer/src/data/sort'
import { removeDuplicates } from '../src/renderer/src/data/removeDuplicates'

const n = Number(process.argv[2] ?? 5000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const baseCols = Object.keys(base.cells).reduce((m, k) => Math.max(m, +k.split(',')[1]), 0) + 1
const FORM = process.argv[3] === 'formulas'
const cols = baseCols + (FORM ? 2 : 0)
const baseRows = Object.keys(base.cells).reduce((m, k) => Math.max(m, +k.split(',')[0]), 0)
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) {
  const sr = r === 0 ? 0 : 1 + ((r - 1) % baseRows)
  for (let c = 0; c < baseCols; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x }
  if (FORM && r > 0) { cells[r + ',' + baseCols] = { raw: `=D${r + 1}&"-"&E${r + 1}` }; cells[r + ',' + (baseCols + 1)] = { raw: `=VLOOKUP(D${r + 1},$D$2:$E$${n},2,FALSE)` } }
}
const store = new WorkbookStore()
store.load({ ...src, sheets: [{ ...base, cells }] }, 'i.xlsx', undefined)
const sh = store.active
const last = n - 1
let fails = 0
const check = (name: string, ok: boolean, detail = ''): void => { if (!ok) fails++; console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')) }

interface Snap { raw: string; val: string; disp: string; style: string }
function snap(rows: number, cs: number): Snap[][] {
  const out: Snap[][] = []
  for (let r = 0; r < rows; r++) {
    const row: Snap[] = []
    for (let c = 0; c < cs; c++) row.push({ raw: store.getRaw(sh, r, c), val: JSON.stringify(store.getValue(sh, r, c)), disp: store.getDisplay(sh, r, c).text, style: JSON.stringify(store.getStyle(sh, r, c) ?? null) })
    out.push(row)
  }
  return out
}
function firstDiff(a: Snap[][], b: Snap[][]): string {
  for (let r = 0; r < Math.max(a.length, b.length); r++) for (let c = 0; c < cols; c++) {
    const x = a[r]?.[c], y = b[r]?.[c]
    if (JSON.stringify(x) !== JSON.stringify(y)) return `r${r} c${c}: ${JSON.stringify(x)} vs ${JSON.stringify(y)}`
  }
  return ''
}
const rows = n
const before = snap(rows, cols)
const hasFormula = before.some((r) => r.some((x) => x.raw.startsWith('=')))
console.log(`${n} rows x ${cols} cols, formulas present: ${hasFormula}`)

const bag = (rs: Snap[][], f: (x: Snap) => string): Map<string, number> => {
  const m = new Map<string, number>()
  rs.forEach((r) => { const k = (FORM ? r.slice(0, -1) : r).map(f).join('\u001f'); m.set(k, (m.get(k) ?? 0) + 1) })
  return m
}
const sameBag = (a: Map<string, number>, b: Map<string, number>): number => { let d = 0; for (const [k, v] of a) if (b.get(k) !== v) d++; return d + (a.size === b.size ? 0 : 1) }

// ---- sort by column 4 (asc) ----
const keyCol = 4
const dataBefore = before.slice(1)
sortRange(store, sh, { r1: 0, c1: 0, r2: last, c2: cols - 1 }, { hasHeaders: true, levels: [{ columnIndex: keyCol, sortOn: 'values', order: 'asc' }] } as never)
const afterSort = snap(rows, cols)
check('sort: header row untouched', JSON.stringify(afterSort[0]) === JSON.stringify(before[0]))
const keysAfter = afterSort.slice(1).map((r) => JSON.parse(r[keyCol].val))
let ordered = true
for (let i = 1; i < keysAfter.length; i++) {
  const a = keysAfter[i - 1], b = keysAfter[i]
  if (a === null || b === null) continue
  if (typeof a === typeof b && a > b) { ordered = false; break }
}
check('sort: key column ascending', ordered)
const f = hasFormula ? (x: Snap): string => x.disp + '|' + x.style : (x: Snap): string => x.raw + '|' + x.style
const dd = sameBag(bag(dataBefore, f), bag(afterSort.slice(1), f))
check('sort: same rows, none lost / duplicated / altered', dd === 0, dd ? `${dd} distinct rows differ` : '')
if (FORM) {
  // the VLOOKUP column must equal the first row (in the new order) whose key matches, as Excel evaluates it
  const first = new Map<string, string>()
  for (let r = 1; r < rows; r++) { const k = afterSort[r][3].disp; if (!first.has(k)) first.set(k, afterSort[r][4].disp) }
  let bad = 0
  for (let r = 1; r < rows; r++) if (afterSort[r][cols - 1].disp !== first.get(afterSort[r][3].disp)) bad++
  check('sort: VLOOKUP column matches first-match semantics', bad === 0, bad ? bad + ' cells' : '')
  let badCat = 0
  for (let r = 1; r < rows; r++) if (afterSort[r][cols - 2].disp !== afterSort[r][3].disp + '-' + afterSort[r][4].disp) badCat++
  check('sort: row-relative formulas follow their rows', badCat === 0, badCat ? badCat + ' cells' : '')
}
store.undo()
const d1 = firstDiff(before, snap(rows, cols))
check('sort: undo restores every cell', d1 === '', d1)
store.redo(); store.undo()
check('sort: redo/undo round trip', firstDiff(before, snap(rows, cols)) === '')

// ---- sort through the AutoFilter menu (a separate code path) ----
store.toggleAutoFilter(sh)
store.sortFilterRange(sh, keyCol, 'desc')
const afterMenu = snap(rows, cols)
const km = afterMenu.slice(1).map((r) => JSON.parse(r[keyCol].val))
let desc = true
for (let i = 1; i < km.length; i++) { const x = km[i - 1], y = km[i]; if (typeof x === 'number' && typeof y === 'number' && x < y) { desc = false; break } }
check('menu sort: key column descending', desc)
const dm = sameBag(bag(dataBefore, f), bag(afterMenu.slice(1), f))
check('menu sort: same rows, none lost / duplicated / altered', dm === 0, dm ? dm + ' distinct rows differ' : '')
store.undo()
const dmu = firstDiff(before, snap(rows, cols))
check('menu sort: undo restores every cell', dmu === '', dmu)
store.toggleAutoFilter(sh)

// ---- remove duplicates on column 6 ----
const dedupeCol = 6
const seen = new Set<string>(); let expectKept = 0
for (let r = 1; r < rows; r++) { const k = before[r][dedupeCol].val.toLowerCase(); if (!seen.has(k)) { seen.add(k); expectKept++ } }
const res = removeDuplicates(store, sh, { r1: 0, c1: 0, r2: last, c2: cols - 1 }, [dedupeCol], true)
check('dedupe: kept count matches model', res.remainingCount === expectKept, `${res.remainingCount} vs ${expectKept}`)
const keptRows = snap(rows, cols)
const keptKeys = new Set<string>(); let dup = false
for (let r = 1; r <= expectKept; r++) { const k = keptRows[r][dedupeCol].val.toLowerCase(); if (keptKeys.has(k)) dup = true; keptKeys.add(k) }
check('dedupe: no duplicates left, first occurrences kept', !dup && [...seen].every((k) => keptKeys.has(k)))
let tailEmpty = true
for (let r = expectKept + 1; r < rows; r++) if (keptRows[r].some((x) => x.raw !== '')) tailEmpty = false
check('dedupe: rows below are cleared', tailEmpty)
store.undo()
const d2 = firstDiff(before, snap(rows, cols))
check('dedupe: undo restores every cell', d2 === '', d2)

// ---- insert / delete rows ----
store.insertRows(sh, 5, 3)
const ins = snap(rows + 3, cols)
let insOk = true
for (let r = 5; r < 8; r++) if (ins[r].some((x) => x.raw !== '')) insOk = false
if (!hasFormula) {
  for (let r = 0; r < 5; r++) if (ins[r].some((x, c) => x.raw !== before[r][c].raw)) insOk = false
  for (let r = 5; r < rows; r++) if (ins[r + 3].some((x, c) => x.raw !== before[r][c].raw)) insOk = false
}
check('insert rows: blank rows, others shifted down intact', insOk)
store.undo()
const d3 = firstDiff(before, snap(rows, cols))
check('insert rows: undo restores', d3 === '', d3)
store.deleteRows(sh, 5, 3)
store.undo()
const d4 = firstDiff(before, snap(rows, cols))
check('delete rows: undo restores', d4 === '', d4)

// ---- filter ----
store.toggleAutoFilter(sh)
const pick = before[1][keyCol].disp
const want = before.map((r, i) => ({ i, v: r[keyCol].disp })).filter((x) => x.i > 0 && x.v === pick).map((x) => x.i)
store.setColumnFilter(sh, keyCol, { selected: [pick] })
const hidden = new Set(sh.layout.autoFilter?.filteredOutRows ?? [])
const visible: number[] = []
for (let r = 1; r < rows; r++) if (!hidden.has(r)) visible.push(r)
check('filter: visible rows == rows matching the value', JSON.stringify(visible) === JSON.stringify(want), `${visible.length} vs ${want.length}`)
store.clearAllFilters(sh); store.toggleAutoFilter(sh)
check('filter: clearing leaves data untouched', firstDiff(before, snap(rows, cols)) === '')

// ---- save and reopen ----
const bytes = await writeXlsx(store.toData())
const s2 = new WorkbookStore(); s2.load(await readXlsx(bytes), 'r.xlsx', undefined)
let rt = ''
for (let r = 0; r < rows && !rt; r++) for (let c = 0; c < cols; c++) {
  const a = before[r][c], rawB = s2.getRaw(s2.active, r, c), dispB = s2.getDisplay(s2.active, r, c).text
  const sameRaw = a.raw === rawB || (a.raw !== '' && Number(a.raw) === Number(rawB) && !isNaN(Number(a.raw)))
  if (!sameRaw || a.disp !== dispB) { rt = `r${r} c${c}: raw ${JSON.stringify(a.raw)} vs ${JSON.stringify(rawB)}, shown ${JSON.stringify(a.disp)} vs ${JSON.stringify(dispB)}`; break }
}
check('save + reopen: raw text and displayed text identical', rt === '', rt)

// ---- paste block, then check nothing else moved ----
const blk = Array.from({ length: 1000 }, (_, i) => ({ r: 1 + i, c: cols + 2, raw: String(i * 1.1) }))
store.setCells(sh, blk, 'paste')
check('paste: values land exactly, rest untouched', blk.every((x) => store.getRaw(sh, x.r, x.c) === x.raw) && firstDiff(before, snap(rows, cols)) === '')
store.undo()
check('paste: undo clears', blk.every((x) => store.getRaw(sh, x.r, x.c) === ''))
console.log(fails ? `\n${fails} FAILED` : '\nALL PASSED')
process.exit(fails ? 1 : 0)

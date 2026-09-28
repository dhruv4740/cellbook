// Stress test on the real despatch workbook (samples/private, gitignored), scaled up by repeating its rows.
//   npx vite-node scripts/perf-private.ts [rowsList e.g. 521,5000,50000,200000]
import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { writeXlsx, readXlsx } from '../src/renderer/src/io/xlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
import { sortRange } from '../src/renderer/src/data/sort'
import { computePivot, PivotDef } from '../src/renderer/src/data/pivot'
import { applySubtotals } from '../src/renderer/src/data/subtotal'
import { removeDuplicates } from '../src/renderer/src/data/removeDuplicates'
import { advancedFilterRows } from '../src/renderer/src/data/advancedFilter'

const targets = (process.argv[2] ?? '521,5000,50000,200000').split(',').map(Number)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const baseRows = Object.keys(base.cells).reduce((m, k) => Math.max(m, +k.split(',')[0]), 0) // last data row index
const cols = Object.keys(base.cells).reduce((m, k) => Math.max(m, +k.split(',')[1]), 0) + 1
const hdr = (c: number): string => String(base.cells['0,' + c]?.raw ?? '')
const srcCol = hdr(4) === 'SOURCE' ? 4 : 4, dstCol = 5

const ms = (t: number): string => (performance.now() - t).toFixed(0).padStart(7) + ' ms'
const line = (label: string, t: number, extra = ''): void => console.log('  ' + label.padEnd(44) + ms(t) + (extra ? '  ' + extra : ''))

function scaled(n: number): typeof src {
  const cells: typeof base.cells = {}
  for (let r = 0; r < n; r++) {
    const sr = r === 0 ? 0 : 1 + ((r - 1) % baseRows)
    for (let c = 0; c < cols; c++) {
      const cell = base.cells[sr + ',' + c]
      if (!cell) continue
      // make the shipment id / GIR numbers distinct per repeat so dedupe/lookups behave like real, unique data
      cells[r + ',' + c] = cell
    }
  }
  return { ...src, sheets: [{ ...base, cells }] }
}

async function run(n: number): Promise<void> {
  console.log(`\n=== ${n.toLocaleString()} rows x ${cols} cols ===`)
  const data = scaled(n)
  let t = performance.now()
  const bytes = await writeXlsx(data)
  line('writeXlsx (build file)', t, (bytes.length / 1048576).toFixed(2) + ' MB')
  t = performance.now()
  const parsed = parseXlsx(bytes)
  line('parseXlsx (fast reader)', t)
  t = performance.now()
  const store = new WorkbookStore()
  store.load(parsed, 'perf.xlsx', undefined)
  line('store.load (HyperFormula build)', t)
  const sh = store.active
  const last = n - 1

  t = performance.now()
  const dstLetter = (c: number): string => String.fromCharCode(65 + c)
  // helper columns: lookups and conditional aggregates over the whole table
  const helper = cols + 1
  const cells: { r: number; c: number; raw: string }[] = []
  for (let r = 1; r <= Math.min(last, 5000); r++) cells.push({ r, c: helper, raw: `=VLOOKUP(${dstLetter(3)}${r + 1},$${dstLetter(3)}$2:$${dstLetter(4)}$${last + 1},2,FALSE)` })
  store.setCells(sh, cells, 'vlookup')
  line(`VLOOKUP x${cells.length} (exact, full-column table)`, t)

  t = performance.now()
  const xl: { r: number; c: number; raw: string }[] = []
  for (let r = 1; r <= Math.min(last, 5000); r++) xl.push({ r, c: helper + 1, raw: `=XLOOKUP(${dstLetter(3)}${r + 1},$${dstLetter(3)}$2:$${dstLetter(3)}$${last + 1},$${dstLetter(9)}$2:$${dstLetter(9)}$${last + 1},"n/a")` })
  store.setCells(sh, xl, 'xlookup')
  line(`XLOOKUP x${xl.length}`, t)

  t = performance.now()
  const sumifs: { r: number; c: number; raw: string }[] = []
  for (let r = 1; r <= Math.min(last, 2000); r++) sumifs.push({ r, c: helper + 2, raw: `=SUMIFS($${dstLetter(9)}$2:$${dstLetter(9)}$${last + 1},$${dstLetter(srcCol)}$2:$${dstLetter(srcCol)}$${last + 1},${dstLetter(srcCol)}${r + 1})` })
  store.setCells(sh, sumifs, 'sumifs')
  line(`SUMIFS x${sumifs.length} by source`, t)

  t = performance.now()
  store.setRaw(sh, 0, helper + 3, `=COUNTIF(${dstLetter(srcCol)}2:${dstLetter(srcCol)}${last + 1},"U815")`)
  line('COUNTIF over whole column', t, String(store.getValue(sh, 0, helper + 3)))

  t = performance.now()
  store.setRaw(sh, 1, 9, '999')
  line('edit one cell (dependents recalc)', t)

  t = performance.now()
  const rg = { r1: 0, c1: 0, r2: last, c2: cols - 1 }
  sortRange(store, sh, rg, { hasHeaders: true, levels: [{ column: srcCol, order: 'asc' } as never] } as never)
  line('sort by SOURCE', t)

  t = performance.now()
  store.toggleAutoFilter(sh)
  line('turn on AutoFilter', t)
  t = performance.now()
  store.setColumnFilter(sh, srcCol, { selected: ['U815'] })
  line('filter SOURCE = U815', t, `hidden ${sh.layout.autoFilter?.filteredOutRows.length} rows`)
  t = performance.now()
  store.setColumnFilter(sh, dstCol, { selected: [store.getDisplay(sh, 1, dstCol).text] })
  line('add filter on DESTINATION', t)
  t = performance.now()
  store.clearAllFilters(sh)
  store.toggleAutoFilter(sh)
  line('clear + remove filter', t)

  t = performance.now()
  const cr = advancedFilterRows([...Array(cols)].map((_, c) => hdr(c)), last, (i, c) => ({ text: store.getDisplay(sh, 1 + i, c).text, value: store.getValue(sh, 1 + i, c) }), [[hdr(srcCol)], ['U815'], ['UPY']])
  line('advanced filter (2 criteria rows)', t, `${cr.length} rows`)

  t = performance.now()
  const def: PivotDef = { id: 'p', sourceSheet: sh.name, source: { r1: 0, c1: 0, r2: last, c2: cols - 1 }, anchor: { r: 0, c: cols + 8 }, rows: [srcCol], cols: [], values: [{ col: 9, func: 'sum' }] }
  const piv = computePivot(store, def)
  line('pivot (rows=SOURCE, sum NET)', t, Array.isArray(piv) ? piv.length + ' output rows' : String(piv))

  t = performance.now()
  removeDuplicates(store, sh, { r1: 0, c1: 0, r2: last, c2: cols - 1 }, [6], true)
  line('remove duplicates on STN NO', t)

  t = performance.now()
  const snap = store.toData()
  line('toData (snapshot for save)', t)
  t = performance.now()
  const out = await writeXlsx(snap)
  line('save: writeXlsx', t, (out.length / 1048576).toFixed(2) + ' MB')
  t = performance.now()
  const back = await readXlsx(out)
  line('reopen: readXlsx', t)
  t = performance.now()
  const s2 = new WorkbookStore()
  s2.load(back, 'perf.xlsx', undefined)
  line('reopen: store.load', t)

  t = performance.now()
  store.insertRows(sh, 5, 3)
  line('insert 3 rows near top', t)
  t = performance.now()
  store.undo()
  line('undo insert', t)
  t = performance.now()
  store.setCells(sh, Array.from({ length: Math.min(last, 20000) }, (_, i) => ({ r: 1 + i, c: cols + 12, raw: String(i) })), 'paste')
  line('paste 20k cells', t)
  t = performance.now()
  store.clear(sh, [{ r1: 1, c1: cols + 12, r2: Math.min(last, 20000) + 1, c2: cols + 12 }])
  line('clear 20k cells', t)
  console.log('  heap', (process.memoryUsage().heapUsed / 1048576).toFixed(0), 'MB')
}

for (const n of targets) await run(n)

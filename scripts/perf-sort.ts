import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
import { sortRange } from '../src/renderer/src/data/sort'
import { removeDuplicates } from '../src/renderer/src/data/removeDuplicates'
const n = Number(process.argv[2] ?? 50000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
const s = new WorkbookStore()
let t = performance.now()
s.load({ ...src, sheets: [{ ...base, cells }] }, 'x', undefined)
console.log('load', (performance.now() - t).toFixed(0))
const sh = s.active
const rg = { r1: 0, c1: 0, r2: n - 1, c2: 20 }
if (process.argv[3] !== 'dedupe') {
  t = performance.now()
  sortRange(s, sh, rg, { hasHeaders: true, levels: [{ columnIndex: 4, sortOn: 'values', order: 'asc' } as never] } as never)
  console.log('sort', (performance.now() - t).toFixed(0))
} else {
  t = performance.now()
  removeDuplicates(s, sh, rg, [6], true)
  console.log('dedupe', (performance.now() - t).toFixed(0))
}

import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
const n = Number(process.argv[2] ?? 5000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
const s = new WorkbookStore()
let t = performance.now()
s.load({ ...src, sheets: [{ ...base, cells }] }, 'x', undefined)
console.log('load', (performance.now() - t).toFixed(0))
const sh = s.active
t = performance.now()
s.setCells(sh, Array.from({ length: n - 1 }, (_, i) => ({ r: i + 1, c: 23, raw: `=VLOOKUP(D${i + 2},$D$2:$E$${n},2,FALSE)` })), 'v')
console.log('vlookup x', n - 1, (performance.now() - t).toFixed(0))
t = performance.now(); s.setRaw(sh, 1, 3, 'ZZ'); console.log('edit key cell', (performance.now() - t).toFixed(0))
t = performance.now(); s.setRaw(sh, 1, 9, '5'); console.log('edit unrelated col', (performance.now() - t).toFixed(0))
console.log(s.getValue(sh, 5, 23))

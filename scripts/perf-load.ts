import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
const n = Number(process.argv[2] ?? 50000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
const s = new WorkbookStore()
const data = { ...src, sheets: [{ ...base, cells }] }
let t = performance.now()
s.load(data, 'x', undefined)
console.log('load', (performance.now() - t).toFixed(0), 'ms; heap', (process.memoryUsage().heapUsed / 1048576).toFixed(0), 'MB')

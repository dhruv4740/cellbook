import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { writeXlsx } from '../src/renderer/src/io/xlsx'
const n = Number(process.argv[2] ?? 20000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
const t = performance.now()
const out = await writeXlsx({ ...src, sheets: [{ ...base, cells }] })
console.log('writeXlsx', n, (performance.now() - t).toFixed(0), 'ms', (out.length / 1048576).toFixed(2), 'MB')

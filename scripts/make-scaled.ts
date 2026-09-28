// Writes samples/private/scaled-<rows>.xlsx by repeating the real despatch workbook's rows (gitignored folder).
import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { writeXlsx } from '../src/renderer/src/io/xlsx'
const n = Number(process.argv[2] ?? 50000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
// a lookup column so the file has formulas like a real working copy
for (let r = 1; r < Math.min(n, 20000); r++) cells[r + ',22'] = { raw: `=VLOOKUP(D${r + 1},$D$2:$E$${n},2,FALSE)` }
const out = await writeXlsx({ ...src, sheets: [{ ...base, cells }] })
fs.writeFileSync(`samples/private/scaled-${n}.xlsx`, out)
console.log('wrote', n, (out.length / 1048576).toFixed(1), 'MB')

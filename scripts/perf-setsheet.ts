import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
const n = Number(process.argv[2] ?? 50000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
const s = new WorkbookStore()
s.load({ ...src, sheets: [{ ...base, cells }] }, 'x', undefined)
const id = s.hfId(s.active)
let t = performance.now()
const m = s.hf.getSheetSerialized(id)
console.log('getSheetSerialized', (performance.now() - t).toFixed(0))
// reverse the data rows (a stand-in for a sort)
const rev = [m[0], ...m.slice(1).reverse()]
t = performance.now()
s.hf.setSheetContent(id, rev as never)
console.log('setSheetContent (whole sheet)', (performance.now() - t).toFixed(0))
t = performance.now()
s.hf.setCellContents({ sheet: id, row: 0, col: 0 }, m as never)
console.log('setCellContents block back', (performance.now() - t).toFixed(0))

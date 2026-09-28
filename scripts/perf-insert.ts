import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
const n = Number(process.argv[2] ?? 5000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
for (let r = 1; r < n; r++) {
  cells[r + ',23'] = { raw: `=VLOOKUP(D${r + 1},$D$2:$E$${n},2,FALSE)` }
  if (r < 2000) cells[r + ',24'] = { raw: `=SUMIFS($J$2:$J$${n},$E$2:$E$${n},E${r + 1})` }
}
const s = new WorkbookStore()
let t = performance.now()
s.load({ ...src, sheets: [{ ...base, cells }] }, 'x', undefined)
console.log('load with formulas', (performance.now() - t).toFixed(0))
t = performance.now(); s.insertRows(s.active, 5, 3); console.log('insert rows', (performance.now() - t).toFixed(0))
t = performance.now(); s.undo(); console.log('undo', (performance.now() - t).toFixed(0))
{
  const sh = s.active
  const id = s.hfId(sh)
  t = performance.now(); const snap = s.hf.getSheetSerialized(id); console.log('getSheetSerialized', (performance.now() - t).toFixed(0), snap.length)
  t = performance.now(); s.hf.addRows(id, [5, 3]); console.log('hf.addRows', (performance.now() - t).toFixed(0))
  t = performance.now(); s.hf.removeRows(id, [5, 3]); console.log('hf.removeRows (inverse)', (performance.now() - t).toFixed(0))
  t = performance.now(); s.hf.setSheetContent(id, snap); console.log('hf.setSheetContent (full restore)', (performance.now() - t).toFixed(0))
}

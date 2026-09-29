// Round trip of the real despatch workbook through Cellbook: parse -> store -> save. rt_private.py compares it with the original in real Excel.
import fs from 'fs'
import { parseXlsx } from '../../src/renderer/src/io/fastXlsx'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'
import { WorkbookStore } from '../../src/renderer/src/model/store'
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const s = new WorkbookStore()
s.load(src, 'x', null)
fs.mkdirSync('scripts/excel-oracle/out', { recursive: true })
fs.writeFileSync('scripts/excel-oracle/out/rt-private.xlsx', await writeXlsx(s.toData()))
// a second file: sorted by SOURCE then a filter applied, to check what Excel sees after Cellbook's edit operations
import { sortRange } from '../../src/renderer/src/data/sort'
const used = s.usedRange(s.active)!
sortRange(s, s.active, { r1: 0, c1: 0, r2: used.r2, c2: used.c2 }, { hasHeaders: true, levels: [{ columnIndex: 4, sortOn: 'values', order: 'asc' } as never] } as never)
s.toggleAutoFilter(s.active)
s.setColumnFilter(s.active, 4, { selected: [s.getDisplay(s.active, 1, 4).text] })
fs.writeFileSync('scripts/excel-oracle/out/rt-private-sorted.xlsx', await writeXlsx(s.toData()))
console.log('written', used.r2 + 1, 'rows')

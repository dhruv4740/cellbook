// Writes a big mixed workbook through Cellbook's fast cell writer; bulk_check.py opens it in real Excel and verifies it.
import fs from 'fs'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'
const cells: Record<string, any> = {}
const styles: any[] = [
  undefined,
  { bold: true, color: '#FF0000', fill: '#FFFF00' },
  { numFmt: '0.00%', hAlign: 'center', vAlign: 'middle', wrap: true },
  { numFmt: '#,##0.00', italic: true, underline: 'double', strike: true, fontFamily: 'Arial', fontSize: 14 },
  { borders: { top: { style: 'thin', color: '#000000' }, bottom: { style: 'double', color: '#0000FF' }, left: { style: 'thick', color: '#00FF00' }, right: { style: 'dashed', color: '#123456' } } },
  { numFmt: 'yyyy-mm-dd', indent: 2, hAlign: 'left' },
  { numFmt: '[$₹-4009] #,##0.00', locked: false, hidden: true },
  { rotation: 90, shrinkToFit: true }
]
const N = 6000
for (let r = 0; r < N; r++) {
  for (let c = 0; c < 6; c++) {
    const k = `${r},${c}`
    const st = styles[(r + c) % styles.length]
    let raw: string
    switch (c) {
      case 0: raw = `Name ${r % 50} & <x>`; break
      case 1: raw = String((r * 7) % 1000 + 0.25); break
      case 2: raw = r % 2 ? 'TRUE' : 'FALSE'; break
      case 3: raw = `=B${r + 1}*2`; break
      case 4: raw = r % 100 === 0 ? "'007" : r % 3 ? '45292' : ''; break
      default: raw = r % 10 === 0 ? `=A${r + 1}&"!"` : ''
    }
    const cell: any = {}
    if (raw !== '') cell.raw = raw
    if (c === 3) cell.value = ((r * 7) % 1000 + 0.25) * 2
    if (c === 5 && raw) cell.value = `Name ${r % 50} & <x>!`
    if (st) cell.style = st
    if (Object.keys(cell).length) cells[k] = cell
  }
}
cells['0,7'] = { raw: '=1/0', value: { error: '#DIV/0!' } }
cells['1,7'] = { raw: '=A1=A1', value: true }
cells['2,7'] = { raw: '  leading and trailing  ' }
cells['3,7'] = { raw: 'line1\nline2' }
const layout: any = { colWidths: { 0: 140 }, rowHeights: { 2: 40 }, hiddenCols: {}, hiddenRows: { 5: true }, merges: [{ r1: 10, c1: 8, r2: 10, c2: 10 }], frozen: { rows: 1, cols: 0 }, showGridlines: true, showHeadings: true, zoom: 100, notes: { '4,0': { text: 'a note', author: 'Me' } }, comments: {} }
cells['10,8'] = { raw: 'merged' }
const out = await writeXlsx({ sheets: [{ name: 'Big', cells, layout }, { name: 'Small', cells: { '0,0': { raw: 'x' } }, layout: undefined as never }] } as never)
fs.writeFileSync('scripts/excel-oracle/out/bulk.xlsx', out)
fs.writeFileSync('scripts/excel-oracle/out/bulk.json', JSON.stringify({ cells, N }))
console.log('written', Object.keys(cells).length, 'cells', out.length, 'bytes')

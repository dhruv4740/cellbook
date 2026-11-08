import { unzipSync } from 'fflate'
import { writeXlsx } from '../src/renderer/src/io/xlsx'
const out = await writeXlsx({ sheets: [{ name: 'S', cells: { '0,0': { raw: 'Hello', style: { bold: true, numFmt: '0.00', fill: '#FF0000' } }, '0,1': { raw: '42' }, '1,0': { raw: '=A1&"!"', value: 'Hello!' } }, layout: { rowHeights: { 0: 30 }, colWidths: {}, hiddenCols: {}, hiddenRows: {}, merges: [], frozen: { rows: 0, cols: 0 }, showGridlines: true, showHeadings: true, zoom: 100, notes: {}, comments: {} } }] } as never)
const f = unzipSync(out)
const d = new TextDecoder()
console.log(d.decode(f['xl/styles.xml']).slice(300, 1500))
console.log(d.decode(f['xl/sharedStrings.xml']))
console.log(d.decode(f['xl/worksheets/sheet1.xml']))

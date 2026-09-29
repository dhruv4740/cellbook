// Writes a workbook with a picture and a chart; pic_check.py opens it in real Excel and lists the shapes.
import fs from 'fs'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const cells: Record<string, any> = { '0,0': { raw: 'Cat' }, '0,1': { raw: 'Val' }, '1,0': { raw: 'a' }, '1,1': { raw: '1' }, '2,0': { raw: 'b' }, '2,1': { raw: '2' }, '3,0': { raw: 'c' }, '3,1': { raw: '3' } }
const layout = {
  pictures: process.env.NOPIC ? [] : [{ id: 'p1', name: 'Logo', src: PNG, anchor: { r: 2, c: 4, dx: 5, dy: 3 }, width: 120, height: 60, alt: 'company logo' }],
  charts: process.env.NOCHART ? [] : [{ id: 'c1', type: 'bar', subtype: 'clusteredColumn', legendPosition: 'right', anchor: { r: 8, c: 0, dx: 0, dy: 0 }, width: 300, height: 200, dataRange: { r1: 0, c1: 0, r2: 3, c2: 1 }, seriesInColumns: true, title: 'T', elements: { axes: true, axisTitles: false, title: true, dataLabels: false, gridlines: true, legend: true } }]
}
fs.mkdirSync('scripts/excel-oracle/out', { recursive: true })
fs.writeFileSync('scripts/excel-oracle/out/pic.xlsx', await writeXlsx({ sheets: [{ name: 'S', cells, layout }] } as any))
console.log('ok')

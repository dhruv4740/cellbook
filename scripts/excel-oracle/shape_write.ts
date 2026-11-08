// Writes shapes / a text box / an arrow; shape_check.py opens the file in real Excel and lists the shapes.
import fs from 'fs'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'
const shapes = [
  { id: 'a', name: 'Rectangle 1', kind: 'rect', anchor: { r: 1, c: 1, dx: 4, dy: 2 }, width: 120, height: 60, fill: '#4472c4', line: '#2f528f', lineWidth: 1, text: 'Hello', color: '#ffffff', fontSize: 14, bold: true, align: 'ctr' },
  { id: 'b', name: 'TextBox 1', kind: 'textbox', anchor: { r: 8, c: 0, dx: 0, dy: 0 }, width: 150, height: 40, text: 'a note', fontSize: 11, color: '#000000' },
  { id: 'c', name: 'Arrow 1', kind: 'arrow', anchor: { r: 3, c: 5, dx: 0, dy: 0 }, width: 100, height: 50, line: '#c00000', lineWidth: 2 },
  { id: 'd', name: 'Oval 1', kind: 'ellipse', anchor: { r: 12, c: 2, dx: 0, dy: 0 }, width: 80, height: 80, fill: '#ed7d31', line: '#000000', lineWidth: 1 }
]
fs.mkdirSync('scripts/excel-oracle/out', { recursive: true })
fs.writeFileSync('scripts/excel-oracle/out/shapes.xlsx', await writeXlsx({ sheets: [{ name: 'S', cells: { '0,0': { raw: 'x' } }, layout: { shapes } }] } as any))
console.log('ok')

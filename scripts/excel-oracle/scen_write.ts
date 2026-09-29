// Writes a workbook with scenarios; scen_check.py opens it in real Excel and lists them.
import fs from 'fs'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'
const cells: Record<string, any> = { '0,0': { raw: '1' }, '1,0': { raw: '2' }, '2,0': { raw: '=A1+A2' } }
const scenarios = [
  { name: 'Low', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['50', '2'], comment: 'low case' },
  { name: 'High', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['200', '9'] }
]
fs.mkdirSync('scripts/excel-oracle/out', { recursive: true })
fs.writeFileSync('scripts/excel-oracle/out/scen.xlsx', await writeXlsx({ sheets: [{ name: 'S', cells, layout: { scenarios } }] } as any))
console.log('ok')

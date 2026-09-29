// Writes a workbook with sparklines; spark_check.py opens it in real Excel and reads SparklineGroups.
import fs from 'fs'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'
const cells: Record<string, any> = {}
for (let r = 0; r < 3; r++) for (let c = 0; c < 5; c++) cells[`${r},${c}`] = { raw: String((c - 2) * (r + 1)) }
const l = { sparklines: [
  { r: 0, c: 5, type: 'line', src: { r1: 0, c1: 0, r2: 0, c2: 4 }, color: '#376092' },
  { r: 1, c: 5, type: 'column', src: { r1: 1, c1: 0, r2: 1, c2: 4 }, color: '#ff0000' },
  { r: 2, c: 5, type: 'winloss', src: { r1: 2, c1: 0, r2: 2, c2: 4 }, color: '#376092' }
] }
fs.mkdirSync('scripts/excel-oracle/out', { recursive: true })
fs.writeFileSync('scripts/excel-oracle/out/spark.xlsx', await writeXlsx({ sheets: [{ name: 'My Sheet', cells, layout: l }] } as any))
console.log('ok')

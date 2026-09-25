// Benchmarks the fast xlsx reader on a large workbook: 100,000 rows x 15 columns, 2 of them
// formulas. Run with `npx vite-node scripts/bench-xlsx.ts` from the repo root.
import ExcelJS from 'exceljs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { readXlsxExceljs } from '../src/renderer/src/io/xlsx'

const ROWS = 100_000
const COLS = 15

async function buildWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Bench')
  const header = []
  for (let c = 0; c < COLS; c++) header.push('Col' + (c + 1))
  ws.addRow(header)
  for (let r = 0; r < ROWS; r++) {
    const row: (string | number | { formula: string })[] = []
    for (let c = 0; c < COLS - 2; c++) {
      row.push(c % 3 === 0 ? r * c + 0.5 : 'value-' + r + '-' + c)
    }
    const excelRow = r + 2
    row.push({ formula: `A${excelRow}+${r}` })
    row.push({ formula: `SUM(A${excelRow}:C${excelRow})` })
    ws.addRow(row)
  }
  const buf = await wb.xlsx.writeBuffer()
  return new Uint8Array(buf as ArrayBuffer)
}

async function main(): Promise<void> {
  process.stdout.write(`Building ${ROWS.toLocaleString()} x ${COLS} workbook...\n`)
  const t0 = performance.now()
  const bytes = await buildWorkbook()
  const t1 = performance.now()
  process.stdout.write(`  build: ${(t1 - t0).toFixed(0)}ms, size: ${(bytes.length / 1024 / 1024).toFixed(2)} MB\n`)

  const t2 = performance.now()
  const fast = parseXlsx(bytes)
  const t3 = performance.now()
  const fastCellCount = Object.values(fast.sheets[0].cells).length
  process.stdout.write(`\nfastXlsx.parseXlsx:   ${(t3 - t2).toFixed(0)}ms  (${fastCellCount.toLocaleString()} cells)\n`)

  const t4 = performance.now()
  const slow = await readXlsxExceljs(bytes)
  const t5 = performance.now()
  const slowCellCount = Object.values(slow.sheets[0].cells).length
  process.stdout.write(`readXlsxExceljs:      ${(t5 - t4).toFixed(0)}ms  (${slowCellCount.toLocaleString()} cells)\n`)

  process.stdout.write(`\nspeedup: ${((t5 - t4) / (t3 - t2)).toFixed(1)}x\n`)
}

void main()

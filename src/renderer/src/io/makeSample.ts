// Builds samples/sample.xlsx: formulas, formats, merges, widths, frozen panes, two sheets.
// Run with: npx tsx src/renderer/src/io/makeSample.ts  (or vitest, see sample.test.ts)
import ExcelJS from 'exceljs'
import { promises as fs } from 'fs'
import { fileURLToPath } from 'url'
import { dirname, join } from 'path'

export async function buildSampleWorkbook(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Excel Clone'

  const budget = wb.addWorksheet('Budget')
  budget.mergeCells('A1:D1')
  budget.getCell('A1').value = 'Quarterly Budget'
  budget.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF217346' } }
  budget.getCell('A1').alignment = { horizontal: 'center' }

  const headers = ['Category', 'Q1', 'Q2', 'Total']
  headers.forEach((h, i) => {
    const cell = budget.getCell(2, i + 1)
    cell.value = h
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF217346' } }
    cell.alignment = { horizontal: 'center' }
  })

  const rows: [string, number, number][] = [
    ['Salaries', 12000, 12500],
    ['Marketing', 3000, 3200],
    ['R&D', 5000, 5400]
  ]
  rows.forEach(([cat, q1, q2], i) => {
    const r = i + 3
    budget.getCell(r, 1).value = cat
    budget.getCell(r, 2).value = q1
    budget.getCell(r, 3).value = q2
    budget.getCell(r, 4).value = { formula: `B${r}+C${r}` }
    budget.getCell(r, 2).numFmt = '$#,##0'
    budget.getCell(r, 3).numFmt = '$#,##0'
    budget.getCell(r, 4).numFmt = '$#,##0'
  })
  const totalRow = rows.length + 3
  budget.getCell(totalRow, 1).value = 'Total'
  budget.getCell(totalRow, 1).font = { bold: true }
  budget.getCell(totalRow, 2).value = { formula: `SUM(B3:B${totalRow - 1})` }
  budget.getCell(totalRow, 3).value = { formula: `SUM(C3:C${totalRow - 1})` }
  budget.getCell(totalRow, 4).value = { formula: `SUM(D3:D${totalRow - 1})` }
  for (let c = 2; c <= 4; c++) {
    budget.getCell(totalRow, c).numFmt = '$#,##0'
    budget.getCell(totalRow, c).font = { bold: true }
    budget.getCell(totalRow, c).border = { top: { style: 'thin' } }
  }

  budget.getColumn(1).width = 18
  budget.getColumn(2).width = 12
  budget.getColumn(3).width = 12
  budget.getColumn(4).width = 12
  budget.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }]

  const notes = wb.addWorksheet('Notes')
  notes.getCell('A1').value = 'Prepared by Finance'
  notes.getCell('A2').value = { formula: 'Budget!D3' }
  notes.getColumn(1).width = 30

  wb.views = [{ x: 0, y: 0, width: 25000, height: 15000, firstSheet: 0, activeTab: 0, visibility: 'visible' }]
  return wb
}

async function main(): Promise<void> {
  const wb = await buildSampleWorkbook()
  const here = dirname(fileURLToPath(import.meta.url))
  const out = join(here, '..', '..', '..', '..', 'samples', 'sample.xlsx')
  await fs.writeFile(out, Buffer.from(await wb.xlsx.writeBuffer()))
  // eslint-disable-next-line no-console
  console.log('Wrote', out)
}

if (import.meta.url === `file://${process.argv[1]}`) void main()

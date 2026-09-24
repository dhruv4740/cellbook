// Differential test: the fast, hand-rolled xlsx reader (readXlsx, backed by fastXlsx.ts) must
// produce the same WorkbookData as the exceljs-based reader (readXlsxExceljs) for the same bytes.
import { describe, it, expect } from 'vitest'
import { readFile } from 'fs/promises'
import { fileURLToPath } from 'url'
import { dirname, join } from 'path'
import ExcelJS from 'exceljs'
import { readXlsx, readXlsxExceljs, writeXlsx } from './xlsx'
import type { WorkbookData } from '../model/store'

const here = dirname(fileURLToPath(import.meta.url))
const samplePath = join(here, '..', '..', '..', '..', 'samples', 'sample.xlsx')

async function expectEqualReads(bytes: Uint8Array): Promise<WorkbookData> {
  const [fast, slow] = await Promise.all([readXlsx(bytes), readXlsxExceljs(bytes)])
  expect(fast).toEqual(slow)
  return fast
}

describe('fastXlsx vs exceljs reader', () => {
  it('matches on samples/sample.xlsx', async () => {
    const bytes = new Uint8Array(await readFile(samplePath))
    const data = await expectEqualReads(bytes)
    expect(data.sheets.length).toBe(2)
    expect(data.sheets[0].cells['2,3'].raw).toBe('=B3+C3')
  })

  it('matches on a workbook covering styles, merges, frozen panes, shared formulas, rich text, dates, booleans, errors and notes', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Kitchen Sink')

    ws.getCell('A1').value = 'plain'
    ws.getCell('A2').value = 3.14159
    ws.getCell('A3').value = true
    ws.getCell('A4').value = false
    ws.getCell('A5').value = { formula: '1/0', result: { error: '#DIV/0!' } } as never
    ws.getCell('A6').value = new Date(Date.UTC(2024, 5, 15))
    ws.getCell('A6').numFmt = 'yyyy-mm-dd'
    ws.getCell('A7').value = {
      richText: [
        { text: 'bold ', font: { bold: true } },
        { text: 'and red', font: { color: { argb: 'FFFF0000' } } }
      ]
    }

    // shared formula: B1 master, B2/B3 followers
    ws.getCell('B1').value = 1
    ws.getCell('B2').value = 2
    ws.getCell('B3').value = 3
    ws.getCell('C1').value = { formula: 'B1*2', result: 2, shareType: 'shared', ref: 'C1:C3' } as never
    ws.getCell('C2').value = { sharedFormula: 'C1', result: 4 } as never
    ws.getCell('C3').value = { sharedFormula: 'C1', result: 6 } as never

    // styles: font, fill, borders, alignment, number format, theme color
    const styled = ws.getCell('D1')
    styled.value = 'styled'
    styled.font = { name: 'Arial', size: 14, bold: true, italic: true, underline: true, strike: true, color: { argb: 'FF112233' } }
    styled.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } }
    styled.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'double', color: { argb: 'FF0000FF' } }
    }
    styled.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true, indent: 2, textRotation: 45 }
    styled.numFmt = '#,##0.00'
    ws.getCell('D2').value = 'themed'
    ws.getCell('D2').font = { color: { theme: 5 } }
    ws.getCell('D3').value = 12345
    ws.getCell('D3').numFmt = '0.00%'
    ws.getCell('D4').protection = { locked: false, hidden: true }
    ws.getCell('D4').value = 'protected'

    // merges + frozen panes + col/row sizing + hidden rows/cols
    ws.mergeCells('E1:F2')
    ws.getColumn(1).width = 22
    ws.getColumn(5).hidden = true
    ws.getRow(1).height = 28
    ws.getRow(9).hidden = true
    ws.getCell('A9').value = 'hidden row'
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }]
    ws.getCell('A10').value = 'below frozen'

    // notes (legacy comments)
    ws.getCell('A1').note = { texts: [{ font: { bold: true }, text: 'Alice:\n' }, { text: 'a note' }] }
    ws.getCell('B1').note = 'plain note, no author'

    // autoFilter
    ws.autoFilter = { from: 'A1', to: 'D1' }

    const buf = new Uint8Array(await wb.xlsx.writeBuffer())
    const data = await expectEqualReads(buf)

    const s = data.sheets[0]
    expect(s.cells['0,0'].raw).toBe('plain')
    expect(s.cells['4,0'].raw).toBe('=1/0') // formula cells always read back as the formula, not the cached error result
    expect(s.cells['6,0'].raw).toBe('bold and red')
    expect(s.cells['0,2'].raw).toBe('=B1*2')
    expect(s.cells['1,2'].raw).toBe('=B2*2')
    expect(s.cells['2,2'].raw).toBe('=B3*2')
    expect(s.cells['0,3'].style?.fontFamily).toBe('Arial')
    expect(s.cells['0,3'].style?.borders?.top?.style).toBe('thin')
    expect(s.cells['0,3'].style?.vAlign).toBe('middle')
    expect(s.cells['3,3'].style?.locked).toBe(false)
    expect(s.cells['0,0'].style?.locked).toBeUndefined()
    expect(s.layout?.merges).toContainEqual({ r1: 0, c1: 4, r2: 1, c2: 5 })
    expect(s.layout?.frozen).toEqual({ rows: 2, cols: 1 })
    expect(s.layout?.hiddenCols?.[4]).toBe(true)
    expect(s.layout?.hiddenRows?.[8]).toBe(true)
    expect(s.layout?.notes?.['0,0']).toEqual({ author: 'Alice', text: 'a note' })
    // readXlsxExceljs never populates autoFilter when reading a file (see fastXlsx.ts comment);
    // this reader matches that, so nothing to assert here beyond the deep-equal above.
  })

  it('matches on a workbook produced by this app\'s own writeXlsx (round trip)', async () => {
    const data: WorkbookData = {
      activeSheet: 0,
      sheets: [
        {
          name: 'RoundTrip',
          cells: {
            '0,0': { raw: 'Hello', style: { bold: true, color: '#FF0000', fill: '#FFFF00' } },
            '1,0': { raw: '=A1&"!"' },
            '2,0': { raw: '45292', style: { numFmt: 'm/d/yyyy' } }
          },
          layout: {
            colWidths: { 0: 120 },
            merges: [{ r1: 5, c1: 0, r2: 5, c2: 2 }],
            frozen: { rows: 1, cols: 1 },
            tabColor: '#FF0000'
          }
        }
      ]
    }
    const bytes = await writeXlsx(data)
    await expectEqualReads(bytes)
  })
})

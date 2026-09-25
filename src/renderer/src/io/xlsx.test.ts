import { unzipSync } from 'fflate'
import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { readXlsx, writeXlsx, colWidthToPx, pxToColWidth, ptToPx, pxToPt } from './xlsx'
import { postProcessXlsx, type VbaProject } from './xlsxPostProcess'
import { BLANK_TOKEN } from '../engine/filter'
import type { WorkbookData } from '../model/store'
import type { CommentThread } from '../model/types'

describe('column width / row height conversions', () => {
  it('round-trips reasonably close', () => {
    for (const px of [40, 48, 64, 80, 100, 120]) {
      const w = pxToColWidth(px)
      const back = colWidthToPx(w)
      expect(Math.abs(back - px)).toBeLessThanOrEqual(1)
    }
  })
  it('row height pt/px are exact inverses within rounding', () => {
    expect(ptToPx(15)).toBe(20)
    expect(pxToPt(20)).toBe(15)
    expect(Math.abs(pxToPt(ptToPx(24)) - 24)).toBeLessThan(0.1)
  })
})

describe('writeXlsx / readXlsx round trip', () => {
  it('preserves values, formulas, styles, merges, widths, heights, frozen panes, two sheets', async () => {
    const data: WorkbookData = {
      activeSheet: 1,
      sheets: [
        {
          name: 'Data',
          cells: {
            '0,0': { raw: 'Hello', style: { bold: true, color: '#FF0000', fill: '#FFFF00' } },
            '0,1': { raw: '42' },
            '0,2': { raw: 'TRUE' },
            '1,0': { raw: '=A1&"!"' },
            '1,1': { raw: '10' },
            '1,2': { raw: '=B1+B2' },
            '2,0': { raw: '45292', style: { numFmt: 'm/d/yyyy' } }, // 2024-01-01
            '3,0': {
              style: {
                borders: { top: { style: 'thin', color: '#000000' }, bottom: { style: 'double', color: '#0000FF' } },
                hAlign: 'center',
                vAlign: 'middle',
                wrap: true,
                numFmt: '0.00%'
              }
            }
          },
          layout: {
            colWidths: { 0: 120, 1: 80 },
            rowHeights: { 0: 30 },
            hiddenCols: { 4: true },
            hiddenRows: { 9: true },
            merges: [{ r1: 5, c1: 0, r2: 5, c2: 2 }],
            frozen: { rows: 1, cols: 1 },
            showGridlines: true,
            showHeadings: true,
            zoom: 100,
            tabColor: '#FF0000'
          }
        },
        { name: 'Sheet2', cells: { '0,0': { raw: 'second sheet' } } }
      ]
    }

    const bytes = await writeXlsx(data)
    expect(bytes.length).toBeGreaterThan(0)
    const back = await readXlsx(bytes)

    expect(back.sheets.length).toBe(2)
    expect(back.sheets[0].name).toBe('Data')
    expect(back.sheets[1].name).toBe('Sheet2')
    expect(back.activeSheet).toBe(1)

    const s0 = back.sheets[0]
    expect(s0.cells['0,0'].raw).toBe('Hello')
    expect(s0.cells['0,0'].style?.bold).toBe(true)
    expect(s0.cells['0,0'].style?.color).toBe('#FF0000')
    expect(s0.cells['0,0'].style?.fill).toBe('#FFFF00')
    expect(s0.cells['0,1'].raw).toBe('42')
    expect(s0.cells['0,2'].raw).toBe('TRUE')
    expect(s0.cells['1,0'].raw).toBe('=A1&"!"')
    expect(s0.cells['1,2'].raw).toBe('=B1+B2')
    expect(s0.cells['2,0'].raw).toBe('45292')
    expect(s0.cells['2,0'].style?.numFmt).toBe('m/d/yyyy')

    expect(s0.cells['3,0'].style?.borders?.top?.style).toBe('thin')
    expect(s0.cells['3,0'].style?.borders?.bottom?.style).toBe('double')
    expect(s0.cells['3,0'].style?.hAlign).toBe('center')
    expect(s0.cells['3,0'].style?.vAlign).toBe('middle')
    expect(s0.cells['3,0'].style?.wrap).toBe(true)

    expect(s0.layout?.colWidths?.[0]).toBeGreaterThan(100)
    expect(s0.layout?.rowHeights?.[0]).toBeGreaterThan(20)
    expect(s0.layout?.hiddenCols?.[4]).toBe(true)
    expect(s0.layout?.hiddenRows?.[9]).toBe(true)
    expect(s0.layout?.merges).toEqual([{ r1: 5, c1: 0, r2: 5, c2: 2 }])
    expect(s0.layout?.frozen).toEqual({ rows: 1, cols: 1 })
    expect(s0.layout?.tabColor).toBe('#FF0000')

    expect(back.sheets[1].cells['0,0'].raw).toBe('second sheet')
  })
})

describe('reading a workbook built directly with exceljs', () => {
  it('reads values, dates, theme colors and shared formulas', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Sheet1')
    ws.getCell('A1').value = 'plain text'
    ws.getCell('A2').value = 3.14159
    ws.getCell('A3').value = new Date(Date.UTC(2024, 0, 1))
    ws.getCell('A3').numFmt = 'yyyy-mm-dd'
    ws.getCell('B1').value = { formula: 'A2*2' }
    ws.getCell('C1').font = { color: { theme: 4 }, bold: true } // accent1 theme color
    ws.getCell('C1').value = 'themed'
    ws.mergeCells('D1:E2')
    ws.getColumn(1).width = 20
    ws.getRow(1).height = 30

    const buf = await wb.xlsx.writeBuffer()
    const data = await readXlsx(new Uint8Array(buf as ArrayBuffer))
    const s = data.sheets[0]

    expect(s.cells['0,0'].raw).toBe('plain text')
    expect(s.cells['1,0'].raw).toBe('3.14159')
    expect(s.cells['2,0'].style?.numFmt).toBe('yyyy-mm-dd')
    // 2024-01-01 as Excel serial (1900 system) is 45292
    expect(Number(s.cells['2,0'].raw)).toBeCloseTo(45292, 0)
    expect(s.cells['0,1'].raw).toBe('=A2*2')
    expect(s.cells['0,2'].style?.color?.toUpperCase()).toBe('#4472C4')
    expect(s.layout?.merges).toContainEqual({ r1: 0, c1: 3, r2: 1, c2: 4 })
  })
})

describe('AutoFilter criteria round trip', () => {
  it('preserves per-column checked values (including blanks) and the filter range', async () => {
    const data: WorkbookData = {
      sheets: [
        {
          name: 'Filtered',
          cells: {
            '0,0': { raw: 'Fruit' }, '0,1': { raw: 'Qty' },
            '1,0': { raw: 'Apple' }, '1,1': { raw: '3' },
            '2,0': { raw: '' }, '2,1': { raw: '5' },
            '3,0': { raw: 'Pear' }, '3,1': { raw: '1' }
          },
          layout: {
            autoFilter: {
              range: { r1: 0, c1: 0, r2: 3, c2: 1 },
              columns: {
                0: { selected: ['Apple', BLANK_TOKEN] }, // absolute col 0 -> colId 0 (range starts at c1=0)
                1: { selected: ['3', '5'] } // absolute col 1 -> colId 1
              },
              filteredOutRows: [3]
            }
          }
        }
      ]
    }

    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    const af = back.sheets[0].layout?.autoFilter
    expect(af?.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 1 })
    expect(af?.columns[0]?.selected.sort()).toEqual(['Apple', BLANK_TOKEN].sort())
    expect(af?.columns[1]?.selected.sort()).toEqual(['3', '5'])
  })

  it('leaves the range-only case (no criteria) as a plain autoFilter ref', async () => {
    const data: WorkbookData = {
      sheets: [{ name: 'S', cells: { '0,0': { raw: 'H' } }, layout: { autoFilter: { range: { r1: 0, c1: 0, r2: 0, c2: 0 }, columns: {}, filteredOutRows: [] } } }]
    }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    expect(back.sheets[0].layout?.autoFilter?.range).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 })
    expect(back.sheets[0].layout?.autoFilter?.columns).toEqual({})
  })
})

describe('threaded comments round trip', () => {
  it('saves a comment thread as a legacy note in the fallback format and reads it back as a thread', async () => {
    const thread: CommentThread = {
      author: 'Alice',
      text: 'Please double-check this total.',
      time: '2024-01-01T00:00:00.000Z',
      replies: [
        { id: 'x', author: 'Bob', text: 'Looks right to me.', time: '2024-01-02T00:00:00.000Z' },
        { id: 'y', author: 'Alice', text: 'Thanks!', time: '2024-01-03T00:00:00.000Z' }
      ]
    }
    const data: WorkbookData = {
      sheets: [{ name: 'S', cells: { '0,0': { raw: '42' } }, layout: { comments: { '0,0': thread } } }]
    }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    const s = back.sheets[0]
    expect(s.layout?.comments?.['0,0']).toEqual({
      author: 'Alice',
      text: 'Please double-check this total.',
      time: '',
      replies: [
        { id: 'r0', author: 'Alice', text: 'Looks right to me.', time: '' },
        { id: 'r1', author: 'Alice', text: 'Thanks!', time: '' }
      ]
    })
    // Not saved as a plain note - a cell can't show both.
    expect(s.layout?.notes?.['0,0']).toBeUndefined()
  })

  it('a comment wins over a plain note at the same cell', async () => {
    const thread: CommentThread = { author: 'Alice', text: 'the comment', time: '', replies: [] }
    const data: WorkbookData = {
      sheets: [
        {
          name: 'S',
          cells: { '0,0': { raw: '1' } },
          layout: { notes: { '0,0': { author: 'Old', text: 'a stale note' } }, comments: { '0,0': thread } }
        }
      ]
    }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    const s = back.sheets[0]
    expect(s.layout?.comments?.['0,0']?.text).toBe('the comment')
    expect(s.layout?.notes?.['0,0']).toBeUndefined()
  })
})

describe('.xlsm VBA project round trip', () => {
  it('survives read -> write -> read', async () => {
    const data: WorkbookData = { sheets: [{ name: 'S', cells: { '0,0': { raw: 'has macros' } } }] }
    const plain = await writeXlsx(data)

    // Simulate an .xlsm file: inject a synthetic vbaProject.bin the same way a real save-to-.xlsm
    // does, since our own writeXlsx (via exceljs) never produces one on its own.
    const bin = new Uint8Array([1, 2, 3, 4, 5, 250, 251, 252])
    const signature = new Uint8Array([9, 9, 9])
    // The signature is only discoverable via its own relationship (xl/_rels/vbaProject.bin.rels),
    // same as in a real workbook - matching postProcessXlsx's "copy the rels part through" approach.
    const signatureRels = new TextEncoder().encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>' +
        '</Relationships>'
    )
    const vba: VbaProject = { bin, signature, signatureRels }
    const asXlsm = postProcessXlsx(plain, [undefined], vba)

    const opened = await readXlsx(asXlsm)
    expect(opened.vbaProject?.bin).toEqual(bin)
    expect(opened.vbaProject?.signature).toEqual(signature)

    // Re-save (as .xlsm - the macros are threaded back through) and read again.
    const resaved = await writeXlsx(opened, opened.vbaProject)
    // Excel only runs (and keeps) macros when the workbook part carries the macro-enabled type.
    const ct = new TextDecoder().decode(unzipSync(resaved)['[Content_Types].xml'])
    expect(ct).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    expect(ct).toContain('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>')
    const reopened = await readXlsx(resaved)
    expect(reopened.vbaProject?.bin).toEqual(bin)
  })

  it('is not written when saving without a vba project (e.g. .xlsx)', async () => {
    const data: WorkbookData = { sheets: [{ name: 'S', cells: {} }] }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    expect(back.vbaProject).toBeUndefined()
  })
})

describe('shared formulas', () => {
  it('resolves followers of a real shared formula range', async () => {
    const ExcelJS = (await import('exceljs')).default
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = 1
    ws.getCell('A2').value = 2
    ws.getCell('B1').value = { formula: 'A1*2', result: 2, shareType: 'shared', ref: 'B1:B2' } as never
    ws.getCell('B2').value = { sharedFormula: 'B1', result: 4 } as never
    const buf = new Uint8Array(await wb.xlsx.writeBuffer())
    const data = await readXlsx(buf)
    expect(data.sheets[0].cells['1,1'].raw).toBe('=A2*2')
  })
})

describe('conditional formatting round trip', () => {
  it('preserves conditional formatting rules in sheet layout', async () => {
    const data: WorkbookData = {
      sheets: [
        {
          name: 'CFSheet',
          cells: {
            '0,0': { raw: '50' },
            '1,0': { raw: '150' }
          },
          layout: {
            conditionalFormats: [
              {
                id: 'cf1',
                priority: 1,
                ranges: [{ r1: 0, c1: 0, r2: 5, c2: 0 }],
                type: 'cellIs',
                operator: 'greaterThan',
                value: '100',
                format: { fill: '#FFC7CE', color: '#9C0006' }
              }
            ]
          }
        }
      ]
    }

    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)

    expect(back.sheets[0].layout?.conditionalFormats).toBeDefined()
    const rules = back.sheets[0].layout!.conditionalFormats!
    expect(rules.length).toBe(1)
    expect(rules[0].type).toBe('cellIs')
    expect(rules[0].operator).toBe('greaterThan')
    expect(rules[0].value).toBe('100')
    expect(rules[0].format?.fill?.toUpperCase()).toBe('#FFC7CE')
    expect(rules[0].format?.color?.toUpperCase()).toBe('#9C0006')
  })
})

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

  it('round-trips data validation covering all allow types, error styles, messages, lists, and dropdown flags', async () => {
    const data: WorkbookData = {
      sheets: [
        {
          name: 'ValidationSheet',
          cells: {},
          layout: {
            dataValidations: [
              // 1. whole number with between, stop error style, input & error messages
              {
                id: 'dv1',
                ranges: [{ r1: 0, c1: 0, r2: 9, c2: 0 }],
                allow: 'whole',
                operator: 'between',
                formula1: '10',
                formula2: '100',
                ignoreBlank: true,
                showInputMessage: true,
                inputTitle: 'Whole Number Hint',
                inputMessage: 'Enter a value between 10 and 100',
                showErrorAlert: true,
                errorStyle: 'stop',
                errorTitle: 'Invalid Number',
                errorMessage: 'Value must be between 10 and 100'
              },
              // 2. decimal with greaterThan, warning error style
              {
                id: 'dv2',
                ranges: [{ r1: 0, c1: 1, r2: 9, c2: 1 }],
                allow: 'decimal',
                operator: 'greaterThan',
                formula1: '0.5',
                errorStyle: 'warning',
                showErrorAlert: true,
                errorTitle: 'Warning',
                errorMessage: 'Should be > 0.5'
              },
              // 3. list with literal items and dropdown enabled (inCellDropdown: true)
              {
                id: 'dv3',
                ranges: [{ r1: 0, c1: 2, r2: 9, c2: 2 }],
                allow: 'list',
                formula1: '"Red,Green,Blue"',
                inCellDropdown: true,
                ignoreBlank: true
              },
              // 4. list with range source and dropdown hidden (inCellDropdown: false)
              {
                id: 'dv4',
                ranges: [{ r1: 0, c1: 3, r2: 9, c2: 3 }],
                allow: 'list',
                formula1: '=Sheet1!$Z$1:$Z$5',
                inCellDropdown: false,
                ignoreBlank: false
              },
              // 5. date with lessThanOrEqual, information error style
              {
                id: 'dv5',
                ranges: [{ r1: 0, c1: 4, r2: 9, c2: 4 }],
                allow: 'date',
                operator: 'lessThanOrEqual',
                formula1: '2025-12-31',
                errorStyle: 'information',
                showErrorAlert: true,
                errorTitle: 'Info',
                errorMessage: 'Date must be within 2025'
              },
              // 6. time with notEqual
              {
                id: 'dv6',
                ranges: [{ r1: 0, c1: 5, r2: 9, c2: 5 }],
                allow: 'time',
                operator: 'notEqual',
                formula1: '12:00:00'
              },
              // 7. textLength with equal
              {
                id: 'dv7',
                ranges: [{ r1: 0, c1: 6, r2: 9, c2: 6 }],
                allow: 'textLength',
                operator: 'equal',
                formula1: '5'
              },
              // 8. custom with formula
              {
                id: 'dv8',
                ranges: [{ r1: 0, c1: 7, r2: 9, c2: 7 }],
                allow: 'custom',
                formula1: '=ISNUMBER(H1)'
              },
              // 9. any allow type with input message only
              {
                id: 'dv9',
                ranges: [{ r1: 0, c1: 8, r2: 9, c2: 8 }],
                allow: 'any',
                showInputMessage: true,
                inputTitle: 'Any value',
                inputMessage: 'Type anything here'
              },
              // 10. multi-range sqref
              {
                id: 'dv10',
                ranges: [
                  { r1: 0, c1: 9, r2: 5, c2: 9 },
                  { r1: 0, c1: 11, r2: 5, c2: 11 }
                ],
                allow: 'whole',
                operator: 'greaterThanOrEqual',
                formula1: '0'
              }
            ]
          }
        }
      ]
    }

    const bytes = await writeXlsx(data)

    // Check raw XML in zip for valid OOXML
    const files = unzipSync(bytes)
    const sheetXml = new TextDecoder('utf-8').decode(files['xl/worksheets/sheet1.xml'])
    expect(sheetXml).toContain('<dataValidations count="10">')
    // Verify showDropDown="1" is emitted ONLY when inCellDropdown === false
    expect(sheetXml).toContain('type="list"')
    expect(sheetXml).toContain('showDropDown="1"')
    // Verify multi-range sqref
    expect(sheetXml).toContain('sqref="J1:J6 L1:L6"')

    // Fast reader round-trip
    const back = await readXlsx(bytes)
    const dvs = back.sheets[0].layout?.dataValidations
    expect(dvs).toBeDefined()
    expect(dvs?.length).toBe(10)

    // Verify 1: whole between
    const r1 = dvs![0]
    expect(r1.allow).toBe('whole')
    expect(r1.operator).toBe('between')
    expect(r1.formula1).toBe('10')
    expect(r1.formula2).toBe('100')
    expect(r1.ignoreBlank).toBe(true)
    expect(r1.showInputMessage).toBe(true)
    expect(r1.inputTitle).toBe('Whole Number Hint')
    expect(r1.inputMessage).toBe('Enter a value between 10 and 100')
    expect(r1.showErrorAlert).toBe(true)
    expect(r1.errorStyle).toBe('stop')
    expect(r1.errorTitle).toBe('Invalid Number')
    expect(r1.errorMessage).toBe('Value must be between 10 and 100')

    // Verify 2: decimal warning
    const r2 = dvs![1]
    expect(r2.allow).toBe('decimal')
    expect(r2.operator).toBe('greaterThan')
    expect(r2.formula1).toBe('0.5')
    expect(r2.errorStyle).toBe('warning')
    expect(r2.showErrorAlert).toBe(true)
    expect(r2.errorTitle).toBe('Warning')

    // Verify 3: list literal with inCellDropdown: true (dropdown shown)
    const r3 = dvs![2]
    expect(r3.allow).toBe('list')
    expect(r3.formula1).toBe('"Red,Green,Blue"')
    expect(r3.inCellDropdown).toBe(true)

    // Verify 4: list range with inCellDropdown: false (dropdown hidden)
    const r4 = dvs![3]
    expect(r4.allow).toBe('list')
    expect(r4.formula1).toBe('=Sheet1!$Z$1:$Z$5')
    expect(r4.inCellDropdown).toBe(false)
    expect(r4.ignoreBlank).toBe(false)

    // Verify 5: date information
    const r5 = dvs![4]
    expect(r5.allow).toBe('date')
    expect(r5.operator).toBe('lessThanOrEqual')
    expect(r5.formula1).toBe('2025-12-31')
    expect(r5.errorStyle).toBe('information')

    // Verify 6: time
    const r6 = dvs![5]
    expect(r6.allow).toBe('time')
    expect(r6.operator).toBe('notEqual')
    expect(r6.formula1).toBe('12:00:00')

    // Verify 7: textLength
    const r7 = dvs![6]
    expect(r7.allow).toBe('textLength')
    expect(r7.operator).toBe('equal')
    expect(r7.formula1).toBe('5')

    // Verify 8: custom
    const r8 = dvs![7]
    expect(r8.allow).toBe('custom')
    expect(r8.formula1).toBe('=ISNUMBER(H1)')

    // Verify 9: any
    const r9 = dvs![8]
    expect(r9.allow).toBe('any')
    expect(r9.showInputMessage).toBe(true)
    expect(r9.inputTitle).toBe('Any value')

    // Verify 10: multi-range sqref
    const r10 = dvs![9]
    expect(r10.ranges.length).toBe(2)
    expect(r10.ranges[0]).toEqual({ r1: 0, c1: 9, r2: 5, c2: 9 })
    expect(r10.ranges[1]).toEqual({ r1: 0, c1: 11, r2: 5, c2: 11 })
  })
})

describe('chart round trip', () => {
  it('round-trips charts covering bar, column, line, area, pie, doughnut, and scatter types', async () => {
    const data: WorkbookData = {
      sheets: [
        {
          name: 'ChartsSheet',
          cells: {
            '0,0': { raw: 'Product' },
            '0,1': { raw: 'Sales' },
            '0,2': { raw: 'Profit' },
            '1,0': { raw: 'Alpha' },
            '1,1': { raw: '200' },
            '1,2': { raw: '50' },
            '2,0': { raw: 'Beta' },
            '2,1': { raw: '300' },
            '2,2': { raw: '80' },
            '3,0': { raw: 'Gamma' },
            '3,1': { raw: '150' },
            '3,2': { raw: '30' }
          },
          layout: {
            charts: [
              // 1. Clustered Column
              {
                id: 'chart1',
                type: 'bar',
                subtype: 'clusteredColumn',
                dataRange: { r1: 0, c1: 0, r2: 3, c2: 2 },
                seriesInColumns: true,
                title: 'Revenue by Product',
                anchor: { r: 5, c: 0, dx: 10, dy: 5 },
                width: 480,
                height: 288,
                legendPosition: 'bottom',
                elements: {
                  title: true,
                  legend: true,
                  gridlines: true,
                  dataLabels: false,
                  axes: true,
                  axisTitles: false
                }
              },
              // 2. Line with markers
              {
                id: 'chart2',
                type: 'line',
                subtype: 'lineMarkers',
                dataRange: { r1: 0, c1: 0, r2: 3, c2: 2 },
                seriesInColumns: true,
                title: 'Growth Trend',
                anchor: { r: 5, c: 8, dx: 0, dy: 0 },
                width: 480,
                height: 288,
                legendPosition: 'right',
                elements: {
                  title: true,
                  legend: true,
                  gridlines: true,
                  dataLabels: false,
                  axes: true,
                  axisTitles: false
                }
              },
              // 3. Pie chart
              {
                id: 'chart3',
                type: 'pie',
                subtype: 'pie',
                dataRange: { r1: 0, c1: 0, r2: 3, c2: 1 },
                seriesInColumns: true,
                title: 'Market Share',
                anchor: { r: 20, c: 0, dx: 0, dy: 0 },
                width: 400,
                height: 250,
                legendPosition: 'right',
                elements: {
                  title: true,
                  legend: true,
                  gridlines: false,
                  dataLabels: false,
                  axes: false,
                  axisTitles: false
                }
              },
              // 4. Doughnut chart
              {
                id: 'chart4',
                type: 'pie',
                subtype: 'doughnut',
                dataRange: { r1: 0, c1: 0, r2: 3, c2: 1 },
                seriesInColumns: true,
                title: 'Doughnut Split',
                anchor: { r: 20, c: 8, dx: 0, dy: 0 },
                width: 400,
                height: 250,
                legendPosition: 'top',
                elements: {
                  title: true,
                  legend: true,
                  gridlines: false,
                  dataLabels: false,
                  axes: false,
                  axisTitles: false
                }
              },
              // 5. Area chart
              {
                id: 'chart5',
                type: 'line',
                subtype: 'area',
                dataRange: { r1: 0, c1: 0, r2: 3, c2: 2 },
                seriesInColumns: true,
                title: 'Volume Area',
                anchor: { r: 35, c: 0, dx: 0, dy: 0 },
                width: 480,
                height: 288,
                legendPosition: 'bottom',
                elements: {
                  title: true,
                  legend: true,
                  gridlines: true,
                  dataLabels: false,
                  axes: true,
                  axisTitles: false
                }
              },
              // 6. Scatter smooth
              {
                id: 'chart6',
                type: 'scatter',
                subtype: 'scatterSmooth',
                dataRange: { r1: 0, c1: 0, r2: 3, c2: 2 },
                seriesInColumns: true,
                title: 'Correlation',
                anchor: { r: 35, c: 8, dx: 0, dy: 0 },
                width: 480,
                height: 288,
                legendPosition: 'left',
                elements: {
                  title: true,
                  legend: true,
                  gridlines: true,
                  dataLabels: false,
                  axes: true,
                  axisTitles: false
                }
              }
            ]
          }
        }
      ]
    }

    const bytes = await writeXlsx(data)

    // Check DrawingML packaging
    const files = unzipSync(bytes)
    expect(files['xl/drawings/drawing1.xml']).toBeDefined()
    expect(files['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined()
    expect(files['xl/charts/chart1.xml']).toBeDefined()
    expect(files['xl/charts/chart6.xml']).toBeDefined()

    const ct = new TextDecoder('utf-8').decode(files['[Content_Types].xml'])
    expect(ct).toContain('ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"')
    expect(ct).toContain('ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"')

    const sheetXml = new TextDecoder('utf-8').decode(files['xl/worksheets/sheet1.xml'])
    expect(sheetXml).toContain('<drawing r:id="')

    // Read back via fast reader
    const back = await readXlsx(bytes)
    const charts = back.sheets[0].layout?.charts
    expect(charts).toBeDefined()
    expect(charts?.length).toBe(6)

    // Verify 1: Column
    const c1 = charts![0]
    expect(c1.type).toBe('bar')
    expect(c1.subtype).toBe('clusteredColumn')
    expect(c1.title).toBe('Revenue by Product')
    expect(c1.legendPosition).toBe('bottom')
    expect(c1.anchor.r).toBe(5)
    expect(c1.anchor.c).toBe(0)
    expect(c1.dataRange).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 })

    // Verify 2: Line with markers
    const c2 = charts![1]
    expect(c2.type).toBe('line')
    expect(c2.subtype).toBe('lineMarkers')
    expect(c2.title).toBe('Growth Trend')
    expect(c2.legendPosition).toBe('right')

    // Verify 3: Pie
    const c3 = charts![2]
    expect(c3.type).toBe('pie')
    expect(c3.subtype).toBe('pie')
    expect(c3.title).toBe('Market Share')

    // Verify 4: Doughnut
    const c4 = charts![3]
    expect(c4.type).toBe('pie')
    expect(c4.subtype).toBe('doughnut')
    expect(c4.title).toBe('Doughnut Split')
    expect(c4.legendPosition).toBe('top')

    // Verify 5: Area
    const c5 = charts![4]
    expect(c5.subtype).toBe('area')
    expect(c5.title).toBe('Volume Area')

    // Verify 6: Scatter
    const c6 = charts![5]
    expect(c6.type).toBe('scatter')
    expect(c6.subtype).toBe('scatterSmooth')
    expect(c6.title).toBe('Correlation')
    expect(c6.legendPosition).toBe('left')
  })

  it('loads charts from samples/chart-sample.xlsx fixture', async () => {
    const { readFile } = await import('fs/promises')
    const { join } = await import('path')
    const bytes = new Uint8Array(await readFile(join(process.cwd(), 'samples', 'chart-sample.xlsx')))
    const wb = await readXlsx(bytes)
    expect(wb.sheets.length).toBe(1)
    const charts = wb.sheets[0].layout?.charts
    expect(charts).toBeDefined()
    expect(charts?.length).toBe(1)
    expect(charts![0].type).toBe('bar')
    expect(charts![0].subtype).toBe('clusteredColumn')
    expect(charts![0].title).toBe('Regional Quarterly Sales')
    expect(charts![0].dataRange).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 })
    expect(charts![0].seriesInColumns).toBe(true)
  })
})

describe('data validation + chart + note on one sheet', () => {
  it('writes worksheet children in OOXML order and reads everything back', async () => {
    const data: WorkbookData = {
      sheets: [
        {
          name: 'Mix',
          cells: { '0,0': { raw: 'A' }, '0,1': { raw: 'B' }, '1,0': { raw: 'x' }, '1,1': { raw: '3' } },
          layout: {
            notes: { '0,0': { text: 'hi', author: 'me' } } as any,
            dataValidations: [{ id: 'd', ranges: [{ r1: 2, c1: 0, r2: 5, c2: 0 }], allow: 'list', formula1: '"a,b"' }],
            charts: [
              {
                id: 'c', type: 'bar', subtype: 'clusteredColumn', dataRange: { r1: 0, c1: 0, r2: 1, c2: 1 },
                seriesInColumns: true, title: 'T', anchor: { r: 3, c: 3, dx: 0, dy: 0 }, width: 300, height: 200,
                legendPosition: 'bottom',
                elements: { title: true, legend: true, gridlines: true, dataLabels: false, axes: true, axisTitles: false }
              }
            ] as any
          }
        }
      ]
    }
    const bytes = await writeXlsx(data)
    const xml = new TextDecoder().decode(unzipSync(bytes)['xl/worksheets/sheet1.xml'])
    const order = ['<dataValidations', '<pageMargins', '<drawing ', '<legacyDrawing ']
      .map((t) => xml.indexOf(t))
      .filter((i) => i !== -1)
    expect(order).toEqual([...order].sort((a, b) => a - b))
    const back = await readXlsx(bytes)
    const l = back.sheets[0].layout!
    expect(l.dataValidations?.length).toBe(1)
    expect(l.charts?.length).toBe(1)
  })
})

describe('defined names', () => {
  it('evaluate in formulas and survive a save', async () => {
    const { store } = await import('../model/store')
    store.load({
      sheets: [
        { name: 'Data', cells: { '0,0': { raw: '10' }, '1,0': { raw: '=TaxRate*A1' }, '2,0': { raw: '=SUM(Amounts)' } } },
        { name: 'Rates', cells: { '0,0': { raw: '1.18' } } }
      ],
      names: [
        { name: 'TaxRate', ref: 'Rates!$A$1' },
        { name: 'Amounts', ref: 'Data!$A$1:$A$2' },
        { name: '_xlnm.Print_Area', ref: 'Data!$A$1:$C$10', sheet: 0 }
      ]
    }, 'n', null)
    expect(store.getDisplay(store.sheets[0], 1, 0).text).toBe('11.8')
    expect(store.getDisplay(store.sheets[0], 2, 0).text).toBe('21.8')
    const back = await readXlsx(await writeXlsx(store.toData()))
    expect(back.names).toEqual([
      { name: 'TaxRate', ref: 'Rates!$A$1' },
      { name: 'Amounts', ref: 'Data!$A$1:$A$2' },
      { name: '_xlnm.Print_Area', ref: 'Data!$A$1:$C$10', sheet: 0 }
    ])
  })
})

describe('data bars, color scales and icon sets', () => {
  it('saves them (used to throw inside ExcelJS) and reads them back with their colors', async () => {
    const data = {
      sheets: [{
        name: 'Sheet1',
        cells: { '0,0': { raw: '1' }, '1,0': { raw: '5' } },
        layout: {
          conditionalFormats: [
            { id: 'a', priority: 1, ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }], type: 'dataBar', dataBar: { color: '#638EC6' } },
            { id: 'b', priority: 2, ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }], type: 'colorScale', colorScale: { minColor: '#F8696B', midColor: '#FFEB84', maxColor: '#63BE7B' } },
            { id: 'c', priority: 3, ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }], type: 'iconSet', iconSet: { name: '3Arrows' } }
          ]
        }
      }]
    } as unknown as WorkbookData
    const back = await readXlsx(await writeXlsx(data))
    const cfs = back.sheets[0].layout!.conditionalFormats!
    expect(cfs.map((c) => c.type)).toEqual(['dataBar', 'colorScale', 'iconSet'])
    expect(cfs[0].dataBar?.color).toBe('#638EC6')
    expect(cfs[1].colorScale).toEqual({ minColor: '#F8696B', midColor: '#FFEB84', maxColor: '#63BE7B' })
    expect(cfs[2].iconSet?.name).toBe('3Arrows')
  })
})

describe('sheet parts after a data bar', () => {
  it('keeps data validations outside the data-bar rule (Excel refused to open such files)', async () => {
    const data = {
      sheets: [{
        name: 'Sheet1',
        cells: { '0,0': { raw: '1' }, '1,0': { raw: '5' } },
        layout: {
          conditionalFormats: [{ id: 'a', priority: 1, ranges: [{ r1: 0, c1: 0, r2: 1, c2: 0 }], type: 'dataBar', dataBar: { color: '#638EC6' } }],
          notes: { '0,0': { text: 'hi', author: 'me' } },
          dataValidations: [{ id: 'v', ranges: [{ r1: 0, c1: 1, r2: 1, c2: 1 }], type: 'list', source: '"Yes,No"', allowBlank: true, showDropdown: true, showErrorMessage: true }]
        }
      }]
    } as unknown as WorkbookData
    const xml = new TextDecoder().decode(unzipSync(await writeXlsx(data))['xl/worksheets/sheet1.xml'])
    const cfEnd = xml.lastIndexOf('</conditionalFormatting>')
    expect(xml.indexOf('<dataValidations')).toBeGreaterThan(cfEnd)
    expect(xml.slice(xml.indexOf('<cfRule'), cfEnd)).not.toContain('<dataValidations')
    expect(xml.indexOf('<legacyDrawing')).toBeGreaterThan(0)
    expect(xml.indexOf('<legacyDrawing')).toBeLessThan(xml.lastIndexOf('<extLst>'))
  })
})

describe('cached formula results', () => {
  it('never writes an error code Excel does not know (Excel refused the file over "#ERROR!")', async () => {
    const data = {
      sheets: [{ name: 'S', cells: { '0,0': { raw: '=1/0', value: { error: '#DIV/0!' } }, '1,0': { raw: '=SUM(T[Qty])', value: { error: '#ERROR!' } } } }]
    } as unknown as WorkbookData
    const xml = new TextDecoder().decode(unzipSync(await writeXlsx(data))['xl/worksheets/sheet1.xml'])
    expect(xml).toContain('#DIV/0!')
    expect(xml).not.toContain('#ERROR!')
  })
})

describe('spilling formulas', () => {
  it('are saved as Excel dynamic arrays (cm="1", ref, metadata part), not as fixed values', async () => {
    const { WorkbookStore } = await import('../model/store')
    const s = new WorkbookStore()
    s.setRaw(s.active, 0, 0, "'00012")
    s.setRaw(s.active, 1, 0, "'00056")
    s.setRaw(s.active, 2, 0, 'x')
    s.setRaw(s.active, 0, 2, '=FILTER(A1:A3,A1:A3<>"x")')
    const files = unzipSync(await writeXlsx(s.toData()))
    const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    expect(xml).toMatch(/<c r="C1"[^>]*cm="1"[^>]*><f t="array" ref="C1:C2">_xlfn\._xlws\.FILTER/)
    expect(xml).not.toContain('r="C2"')
    expect(files['xl/metadata.xml']).toBeDefined()
    expect(new TextDecoder().decode(files['[Content_Types].xml'])).toContain('/xl/metadata.xml')
  })
})

describe('threaded comments written by real Excel', () => {
  it('keeps the text when the fallback note uses Windows line endings', async () => {
    const { legacyNoteToThread } = await import('./xlsxShared')
    const text = '[Threaded comment]\r\n\r\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\r\n\r\nComment:\r\n    Threaded comment'
    expect(legacyNoteToThread({ author: 'tc={6FCF}', text })?.text).toBe('Threaded comment')
  })
})

describe('sparklines', () => {
  it('round-trip through the x14 extension', async () => {
    const defs = [
      { r: 0, c: 5, type: 'line' as const, src: { r1: 0, c1: 0, r2: 0, c2: 4 }, color: '#376092' },
      { r: 1, c: 5, type: 'column' as const, src: { r1: 1, c1: 0, r2: 1, c2: 4 }, color: '#ff0000' },
      { r: 2, c: 5, type: 'winloss' as const, src: { r1: 2, c1: 0, r2: 2, c2: 4 }, color: '#376092' }
    ]
    const cells: Record<string, { raw: string }> = {}
    for (let r = 0; r < 3; r++) for (let c = 0; c < 5; c++) cells[`${r},${c}`] = { raw: String(c - 2 + r) }
    const back = await readXlsx(await writeXlsx({ sheets: [{ name: 'My Sheet', cells: cells as never, layout: { sparklines: defs } as never }] }))
    expect(back.sheets[0].layout?.sparklines).toEqual(expect.arrayContaining(defs))
    expect(back.sheets[0].layout?.sparklines).toHaveLength(3)
  })
})

describe('pictures', () => {
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  it('survive a save and reopen with position, size and bytes', async () => {
    const pic = { id: 'p1', name: 'Logo', src: PNG, anchor: { r: 2, c: 1, dx: 5, dy: 3 }, width: 120, height: 60, alt: 'company logo' }
    const buf = await writeXlsx({ sheets: [{ name: 'S', cells: { '0,0': { raw: 'x' } } as never, layout: { pictures: [pic] } as never }] })
    const back = await readXlsx(buf)
    const got = back.sheets[0].layout?.pictures
    expect(got).toHaveLength(1)
    expect(got![0]).toMatchObject({ name: 'Logo', anchor: { r: 2, c: 1, dx: 5, dy: 3 }, width: 120, height: 60, alt: 'company logo', src: PNG })
  })
  it('coexist with a chart on the same sheet', async () => {
    const pic = { id: 'p1', name: 'Logo', src: PNG, anchor: { r: 0, c: 8, dx: 0, dy: 0 }, width: 50, height: 50 }
    const chart = { id: 'c1', type: 'bar', subtype: 'clusteredColumn', legendPosition: 'right', anchor: { r: 5, c: 0, dx: 0, dy: 0 }, width: 300, height: 200, dataRange: { r1: 0, c1: 0, r2: 2, c2: 1 }, seriesInColumns: true, title: '', elements: { axes: true, axisTitles: false, title: false, dataLabels: false, gridlines: true, legend: true } }
    const cells: Record<string, { raw: string }> = { '0,0': { raw: 'a' }, '0,1': { raw: '1' }, '1,0': { raw: 'b' }, '1,1': { raw: '2' } }
    const back = await readXlsx(await writeXlsx({ sheets: [{ name: 'S', cells: cells as never, layout: { pictures: [pic], charts: [chart] } as never }] }))
    expect(back.sheets[0].layout?.pictures).toHaveLength(1)
    expect(back.sheets[0].layout?.charts).toHaveLength(1)
  })
})

describe('scenarios', () => {
  it('survive a save and reopen', async () => {
    const scenarios = [
      { name: 'Low & "slow"', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['50', 'a<b'], comment: 'bad case' },
      { name: 'High', cells: [{ r: 0, c: 0 }], values: ['200'] }
    ]
    const back = await readXlsx(await writeXlsx({ sheets: [{ name: 'S', cells: { '0,0': { raw: '1' }, '1,0': { raw: '2' } } as never, layout: { scenarios } as never }] }))
    expect(back.sheets[0].layout?.scenarios).toEqual(scenarios)
  })
})

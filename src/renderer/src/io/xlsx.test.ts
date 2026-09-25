import { unzipSync } from 'fflate'
import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { readXlsx, readXlsxExceljs, writeXlsx, colWidthToPx, pxToColWidth, ptToPx, pxToPt } from './xlsx'
import { postProcessXlsx, type VbaProject } from './xlsxPostProcess'
import { BLANK_TOKEN } from '../engine/filter'
import type { WorkbookData } from '../model/store'
import type { CfRule, CommentThread, DataValidationRule } from '../model/types'

async function expectEqualReads(bytes: Uint8Array): Promise<WorkbookData> {
  const [fast, slow] = await Promise.all([readXlsx(bytes), readXlsxExceljs(bytes)])
  expect(fast).toEqual(slow)
  return fast
}

// Both readers mint their own random `id` for each parsed CfRule/DataValidationRule, so comparing
// them needs those ids normalized away first. Scoped to conditionalFormats/dataValidations only
// (rather than the whole WorkbookData) so these tests stay about CF/DV and don't couple to other
// sheet-layout fields (page setup, VBA, ...) owned elsewhere in the codebase.
function cfDvOnly(data: WorkbookData): { conditionalFormats?: CfRule[]; dataValidations?: DataValidationRule[] }[] {
  return data.sheets.map((sd) => ({
    conditionalFormats: sd.layout?.conditionalFormats?.map((r) => ({ ...r, id: '_' })),
    dataValidations: sd.layout?.dataValidations?.map((r) => ({ ...r, id: '_' }))
  }))
}

async function expectEqualCfDvReads(bytes: Uint8Array): Promise<WorkbookData> {
  const [fast, slow] = await Promise.all([readXlsx(bytes), readXlsxExceljs(bytes)])
  expect(cfDvOnly(fast)).toEqual(cfDvOnly(slow))
  return fast
}

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

  it('round-trips every rule type through both readers (fastXlsx and exceljs)', async () => {
    const ranges = [{ r1: 0, c1: 0, r2: 20, c2: 0 }]
    const rules: Omit<CfRule, 'id'>[] = [
      { priority: 1, ranges, type: 'cellIs', operator: 'between', value: '10', value2: '20', format: { fill: '#FFC7CE', color: '#9C0006', bold: true } },
      { priority: 2, ranges, type: 'textContains', text: 'urgent', format: { fill: '#FFEB9C', italic: true } },
      { priority: 3, ranges, type: 'dateOccurring', date: 'lastWeek', format: { color: '#FF0000' } },
      { priority: 4, ranges, type: 'duplicateValues', format: { fill: '#00FF00' } },
      { priority: 5, ranges, type: 'uniqueValues', format: { fill: '#00FFFF' }, stopIfTrue: true },
      { priority: 6, ranges, type: 'top10', bottom: true, percent: true, rank: 25, format: { fill: '#CCCCCC' } },
      { priority: 7, ranges, type: 'aboveAverage', above: false, equalOrAbove: true, format: { strike: true } },
      { priority: 8, ranges, type: 'expression', formula: '=$A1>5', format: { underline: 'single' } },
      { priority: 9, ranges, type: 'dataBar', dataBar: { color: '#638EC6', min: 0, max: 100, showValue: false } },
      { priority: 10, ranges, type: 'colorScale', colorScale: { minColor: '#F8696B', midColor: '#FFEB84', maxColor: '#63BE7B' } },
      { priority: 11, ranges, type: 'iconSet', iconSet: { name: '3Arrows', reverse: true, showValue: false } }
    ]
    const data: WorkbookData = {
      sheets: [
        {
          name: 'AllCf',
          cells: { '0,0': { raw: '15' } },
          layout: { conditionalFormats: rules.map((r, i) => ({ ...r, id: 'seed' + i })) }
        }
      ]
    }

    const bytes = await writeXlsx(data)
    const fastBack = await expectEqualCfDvReads(bytes)
    const cfs = fastBack.sheets[0].layout!.conditionalFormats!
    expect(cfs.length).toBe(rules.length)
    const byType = Object.fromEntries(cfs.map((r) => [r.type, r]))

    expect(byType.cellIs.operator).toBe('between')
    expect(byType.cellIs.value).toBe('10')
    expect(byType.cellIs.value2).toBe('20')
    expect(byType.cellIs.format?.bold).toBe(true)

    expect(byType.textContains.text).toBe('urgent')
    expect(byType.textContains.format?.italic).toBe(true)

    expect(byType.dateOccurring.date).toBe('lastWeek')

    expect(byType.duplicateValues.format?.fill?.toUpperCase()).toBe('#00FF00')
    expect(byType.uniqueValues.stopIfTrue).toBe(true)

    expect(byType.top10.bottom).toBe(true)
    expect(byType.top10.percent).toBe(true)
    expect(byType.top10.rank).toBe(25)

    expect(byType.aboveAverage.above).toBe(false)
    expect(byType.aboveAverage.equalOrAbove).toBe(true)

    expect(byType.expression.formula).toBe('=$A1>5')

    expect(byType.dataBar.dataBar?.color.toUpperCase()).toBe('#638EC6')
    expect(byType.dataBar.dataBar?.min).toBe(0)
    expect(byType.dataBar.dataBar?.max).toBe(100)
    expect(byType.dataBar.dataBar?.showValue).toBe(false)

    expect(byType.colorScale.colorScale?.minColor.toUpperCase()).toBe('#F8696B')
    expect(byType.colorScale.colorScale?.midColor?.toUpperCase()).toBe('#FFEB84')
    expect(byType.colorScale.colorScale?.maxColor.toUpperCase()).toBe('#63BE7B')

    expect(byType.iconSet.iconSet?.name).toBe('3Arrows')
    expect(byType.iconSet.iconSet?.reverse).toBe(true)
    expect(byType.iconSet.iconSet?.showValue).toBe(false)

    // Priorities and per-rule ranges/stopIfTrue all survive too.
    for (const rule of cfs) {
      expect(rule.ranges).toEqual(ranges)
    }
  })

  it('round-trips a Solid Fill data bar via the x14 extLst extension, in both readers', async () => {
    const ranges = [{ r1: 0, c1: 0, r2: 5, c2: 0 }]
    const data: WorkbookData = {
      sheets: [
        {
          name: 'Solid',
          cells: { '0,0': { raw: '10' } },
          layout: {
            conditionalFormats: [
              { id: 'seed', priority: 1, ranges, type: 'dataBar', dataBar: { color: '#0070C0', gradient: false } }
            ]
          }
        }
      ]
    }

    const bytes = await writeXlsx(data)

    // The legacy <cfRule> gets an <extLst> pointing (by GUID) at a worksheet-level x14 extLst that
    // actually carries gradient="0" - assert the raw XML shape the coordinator asked for exactly.
    const files = unzipSync(bytes)
    const sheetXml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    expect(sheetXml).toContain('uri="{B025F937-C7B1-47D3-B67F-A62EEFF666E3}"')
    expect(sheetXml).toContain('<x14:id>')
    expect(sheetXml).toContain('uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"')
    expect(sheetXml).toContain('<x14:conditionalFormattings>')
    expect(sheetXml).toContain('<x14:cfRule type="dataBar" id="')
    expect(sheetXml).toMatch(/<x14:dataBar[^>]*gradient="0"/)
    expect(sheetXml).toContain('<xm:sqref>A1:A6</xm:sqref>')
    // The x14 extLst must be the worksheet's very last child (after pageMargins/pageSetup etc) -
    // the *last* <extLst> in the document, since the legacy cfRule has its own earlier one too.
    expect(sheetXml.lastIndexOf('<extLst>')).toBeGreaterThan(sheetXml.indexOf('<pageMargins'))
    expect(sheetXml.trim().endsWith('</worksheet>')).toBe(true)
    expect(sheetXml).toMatch(/<worksheet\b[^>]*xmlns:x14="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2009\/9\/main"/)
    expect(sheetXml).toMatch(/<worksheet\b[^>]*mc:Ignorable="[^"]*\bx14\b/)
    // The GUID referenced by the legacy cfRule's <x14:id> matches the x14:cfRule's own id attribute.
    const idM = /<x14:id>(\{[0-9A-Fa-f-]+\})<\/x14:id>/.exec(sheetXml)
    expect(idM).toBeTruthy()
    expect(sheetXml).toContain(`<x14:cfRule type="dataBar" id="${idM![1]}">`)

    const fastBack = await expectEqualCfDvReads(bytes)
    const rule = fastBack.sheets[0].layout!.conditionalFormats![0]
    expect(rule.type).toBe('dataBar')
    expect(rule.dataBar?.gradient).toBe(false)
    expect(rule.dataBar?.color.toUpperCase()).toBe('#0070C0')

    // A plain (gradient) data bar - no x14 extension - still reads back as gradient: true.
    const gradientData: WorkbookData = {
      sheets: [
        {
          name: 'Gradient',
          cells: {},
          layout: {
            conditionalFormats: [{ id: 'g', priority: 1, ranges, type: 'dataBar', dataBar: { color: '#638EC6', gradient: true } }]
          }
        }
      ]
    }
    const gradientBytes = await writeXlsx(gradientData)
    const gradientFiles = unzipSync(gradientBytes)
    expect(new TextDecoder().decode(gradientFiles['xl/worksheets/sheet1.xml'])).not.toContain('extLst')
    const gradientBack = await expectEqualCfDvReads(gradientBytes)
    expect(gradientBack.sheets[0].layout!.conditionalFormats![0].dataBar?.gradient).toBe(true)
  })
})

describe('data validation round trip', () => {
  it('round-trips every allow type, including the showDropDown inversion, through both readers', async () => {
    const rules: Omit<DataValidationRule, 'id'>[] = [
      {
        ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }],
        allow: 'list',
        formula1: 'Yes,No,Maybe',
        ignoreBlank: true,
        inCellDropdown: true
      },
      {
        // "In-cell dropdown" unchecked - OOXML writes showDropDown="1" to HIDE it.
        ranges: [{ r1: 1, c1: 0, r2: 1, c2: 0 }],
        allow: 'list',
        formula1: '=$A$1:$A$3',
        ignoreBlank: false,
        inCellDropdown: false
      },
      {
        ranges: [{ r1: 2, c1: 0, r2: 2, c2: 0 }],
        allow: 'whole',
        operator: 'between',
        formula1: '1',
        formula2: '10',
        showInputMessage: true,
        inputTitle: 'Range',
        inputMessage: 'Enter 1-10',
        showErrorAlert: true,
        errorStyle: 'stop',
        errorTitle: 'Bad value',
        errorMessage: 'Must be between 1 and 10.'
      },
      {
        ranges: [{ r1: 3, c1: 0, r2: 3, c2: 0 }],
        allow: 'decimal',
        operator: 'greaterThanOrEqual',
        formula1: '0.5',
        showErrorAlert: true,
        errorStyle: 'warning'
      },
      {
        ranges: [{ r1: 4, c1: 0, r2: 4, c2: 0 }],
        allow: 'date',
        operator: 'lessThan',
        formula1: '45292'
      },
      {
        ranges: [{ r1: 5, c1: 0, r2: 5, c2: 0 }],
        allow: 'textLength',
        operator: 'lessThanOrEqual',
        formula1: '10'
      },
      {
        ranges: [{ r1: 6, c1: 0, r2: 6, c2: 0 }],
        allow: 'custom',
        formula1: '=MOD(A7,2)=0'
      }
    ]
    const data: WorkbookData = {
      sheets: [
        {
          name: 'DV',
          cells: {},
          layout: { dataValidations: rules.map((r, i) => ({ ...r, id: 'seed' + i })) }
        }
      ]
    }

    const bytes = await writeXlsx(data)
    const fastBack = await expectEqualCfDvReads(bytes)
    const dvs = fastBack.sheets[0].layout!.dataValidations!
    expect(dvs.length).toBe(rules.length)
    const byRange = Object.fromEntries(dvs.map((r) => [r.ranges[0].r1, r]))

    expect(byRange[0].allow).toBe('list')
    expect(byRange[0].formula1).toBe('Yes,No,Maybe')
    expect(byRange[0].inCellDropdown).toBe(true)

    expect(byRange[1].allow).toBe('list')
    expect(byRange[1].formula1).toBe('=$A$1:$A$3')
    expect(byRange[1].ignoreBlank).toBe(false)
    expect(byRange[1].inCellDropdown).toBe(false) // showDropDown="1" survived and was un-inverted

    expect(byRange[2].allow).toBe('whole')
    expect(byRange[2].operator).toBe('between')
    expect(byRange[2].formula1).toBe('1')
    expect(byRange[2].formula2).toBe('10')
    expect(byRange[2].inputTitle).toBe('Range')
    expect(byRange[2].inputMessage).toBe('Enter 1-10')
    expect(byRange[2].errorTitle).toBe('Bad value')
    expect(byRange[2].errorMessage).toBe('Must be between 1 and 10.')
    expect(byRange[2].errorStyle).toBe('stop')

    expect(byRange[3].allow).toBe('decimal')
    expect(byRange[3].operator).toBe('greaterThanOrEqual')
    expect(byRange[3].formula1).toBe('0.5')
    expect(byRange[3].errorStyle).toBe('warning')

    expect(byRange[4].allow).toBe('date')
    expect(byRange[4].operator).toBe('lessThan')
    expect(byRange[4].formula1).toBe('45292')

    expect(byRange[5].allow).toBe('textLength')
    expect(byRange[5].formula1).toBe('10')

    expect(byRange[6].allow).toBe('custom')
    expect(byRange[6].formula1).toBe('=MOD(A7,2)=0')
  })

  it('raw XML actually carries showDropDown="1" when the in-cell dropdown is hidden', async () => {
    const data: WorkbookData = {
      sheets: [
        {
          name: 'S',
          cells: {},
          layout: {
            dataValidations: [
              { id: 'dv1', ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], allow: 'list', formula1: 'A,B', inCellDropdown: false }
            ]
          }
        }
      ]
    }
    const bytes = await writeXlsx(data)
    const files = unzipSync(bytes)
    const sheetXml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    expect(sheetXml).toContain('showDropDown="1"')
  })
})

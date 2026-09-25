// Tests the BIFF8 (.xls) reader against small real-world fixtures (Apache POI's public test-data
// corpus; see fixtures/ - each file is a genuine Excel 97-2003 workbook, not a synthetic one).
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { parseXls } from './biff8'
import { isCfbMagic } from '../xlsxShared'

function load(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(__dirname, 'fixtures', name)))
}

describe('isCfbMagic', () => {
  it('recognizes the OLE2/CFB header regardless of extension', () => {
    expect(isCfbMagic(load('Simple.xls'))).toBe(true)
    expect(isCfbMagic(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false) // a .xlsx zip
    expect(isCfbMagic(new Uint8Array([1, 2, 3]))).toBe(false)
  })
})

describe('parseXls', () => {
  it('reads sheet names, text and styled cells (SampleSS.xls)', () => {
    const wb = parseXls(load('SampleSS.xls'))
    expect(wb.sheets.map((s) => s.name)).toEqual(['First Sheet', 'Sheet Number 2', 'Sheet3'])

    const first = wb.sheets[0]
    expect(first.cells['0,0']?.raw).toBe('Test spreadsheet')
    expect(first.cells['1,1']?.raw).toBe('2nd row 2nd column')
    // "This one is red"
    expect(first.cells['3,0']?.raw).toBe('This one is red')
    expect(first.cells['3,0']?.style?.color).toBe('#FF0000')

    const second = wb.sheets[1]
    // "I'm in bold blue, on a yellow background"
    const styled = second.cells['3,0']
    expect(styled?.style?.bold).toBe(true)
    expect(styled?.style?.color).toBe('#003366')
    expect(styled?.style?.fill).toBe('#FFFF00')

    // A SUM formula decodes to real formula text, not just its cached value.
    const formulaCell = second.cells['6,3']
    expect(formulaCell?.raw).toMatch(/^=SUM\(/)
  })

  it('decodes number formats, including custom/escaped ones (DateFormats.xls, Formatting.xls)', () => {
    const wb = parseXls(load('DateFormats.xls'))
    const sheet = wb.sheets[0]
    expect(sheet.cells['0,1']?.style?.numFmt).toBe('d-mmm-yy')
    expect(sheet.cells['1,1']?.style?.numFmt).toBe('m/d/yy;@')
    // The serial number itself is preserved (1900 date system - not adjusted).
    expect(sheet.cells['0,1']?.raw).toBe('39304')

    const wb2 = parseXls(load('Formatting.xls'))
    const s2 = wb2.sheets[0]
    expect(s2.cells['3,1']?.style?.numFmt).toBe('yyyy\\-mm\\-dd')
    expect(s2.cells['12,1']?.style?.numFmt).toBe('"£"#,##0.00')
  })

  it('reads legacy cell notes with author (SimpleWithComments.xls)', () => {
    const wb = parseXls(load('SimpleWithComments.xls'))
    const notes = wb.sheets[0].layout?.notes ?? {}
    expect(notes['0,1']?.author).toBe('Yegor Kozlov')
    expect(notes['0,1']?.text).toContain('first cell')
    expect(notes['1,1']?.text).toContain('second cell')
  })

  it('decodes a shared formula (SharedFormulaTest.xls)', () => {
    const wb = parseXls(load('SharedFormulaTest.xls'))
    const sheet = wb.sheets[0]
    expect(sheet.cells['0,1']?.raw).toBe('=$A1*2')
  })

  it('falls back to the cached value for a formula it cannot fully decode, rather than crashing', () => {
    // 47251.xls and 49219-style large/odd-shaped POI regression fixtures exercise less common
    // record combinations; the important invariant is "never throw, always produce a value".
    const wb = parseXls(load('47251.xls'))
    expect(wb.sheets.length).toBeGreaterThan(0)
  })

  it('parses every bundled fixture without throwing and with at least one sheet', () => {
    const names = [
      'Simple.xls', 'SimpleMultiCell.xls', 'SimpleWithFormula.xls', 'StringFormulas.xls',
      'DateFormats.xls', 'Formatting.xls', 'SampleSS.xls', 'SharedFormulaTest.xls',
      'SimpleWithComments.xls', '47251.xls'
    ]
    for (const name of names) {
      const wb = parseXls(load(name))
      expect(wb.sheets.length).toBeGreaterThan(0)
      expect(wb.activeSheet).toBe(0)
    }
  })

  it('throws a clear error for something that is not a CFB/BIFF8 file', () => {
    expect(() => parseXls(new Uint8Array([1, 2, 3, 4]))).toThrow()
  })
})

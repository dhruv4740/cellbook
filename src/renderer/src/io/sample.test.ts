import { describe, it, expect } from 'vitest'
import { promises as fs } from 'fs'
import path from 'path'
import { buildSampleWorkbook } from './makeSample'
import { readXlsx } from './xlsx'

const SAMPLE_PATH = path.resolve(__dirname, '../../../../samples/sample.xlsx')

describe('samples/sample.xlsx', () => {
  it('builds and writes the sample workbook, and it loads back correctly', async () => {
    const wb = await buildSampleWorkbook()
    const buf = Buffer.from(await wb.xlsx.writeBuffer())
    await fs.mkdir(path.dirname(SAMPLE_PATH), { recursive: true })
    await fs.writeFile(SAMPLE_PATH, buf)

    const bytes = await fs.readFile(SAMPLE_PATH)
    const data = await readXlsx(new Uint8Array(bytes))

    expect(data.sheets.map((s) => s.name)).toEqual(['Budget', 'Notes'])
    const budget = data.sheets[0]
    expect(budget.cells['0,0'].raw).toBe('Quarterly Budget')
    expect(budget.cells['1,0'].raw).toBe('Category')
    expect(budget.cells['3,3'].raw).toBe('=B4+C4') // Total column formula, row 4 (0-indexed row 3)
    expect(budget.layout?.frozen).toEqual({ rows: 2, cols: 0 })
    expect(budget.layout?.merges).toContainEqual({ r1: 0, c1: 0, r2: 0, c2: 3 })
    expect(budget.layout?.colWidths?.[0]).toBeGreaterThan(80)

    const notes = data.sheets[1]
    expect(notes.cells['1,0'].raw).toBe('=Budget!D3')
  })
})

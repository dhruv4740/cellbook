// Round-trip test for Excel Tables (ListObjects): xl/tables/tableN.xml, tableStyleInfo, autoFilter
// and tableColumns (incl. totalsRowFunction/calculatedColumnFormula) - written by
// xlsxTables.ts/xlsxPostProcess.ts, read back by the fast reader (io/fastXlsx.ts).
import { describe, it, expect } from 'vitest'
import { readXlsx, writeXlsx } from './xlsx'
import type { WorkbookData } from '../model/store'
import type { Table } from '../model/types'
import { defaultTableStyleOptions } from '../model/types'

function makeTable(overrides: Partial<Table> = {}): Table {
  return {
    id: 't1',
    name: 'Table1',
    range: { r1: 0, c1: 0, r2: 3, c2: 2 },
    styleName: 'TableStyleMedium9',
    options: { ...defaultTableStyleOptions(), totalRow: true, bandedColumns: true, firstColumn: true },
    columns: [
      { name: 'Item' },
      { name: 'Qty', totalsRowFunction: 'sum' },
      { name: 'Total', calculatedFormula: '=B2*2' }
    ],
    ...overrides
  }
}

describe('table xlsx round trip', () => {
  it('round-trips range, style, options, column names and totals row function', async () => {
    const t = makeTable()
    const data: WorkbookData = {
      sheets: [{
        name: 'S',
        cells: {
          '0,0': { raw: 'Item' }, '0,1': { raw: 'Qty' }, '0,2': { raw: 'Total' },
          '1,0': { raw: 'A' }, '1,1': { raw: '2' }, '1,2': { raw: '=B2*2' },
          '2,0': { raw: 'B' }, '2,1': { raw: '3' }, '2,2': { raw: '=B3*2' },
          '3,1': { raw: '=SUBTOTAL(109,B2:B3)' }
        },
        layout: { tables: [t], autoFilter: { range: { r1: 0, c1: 0, r2: 2, c2: 2 }, columns: {}, filteredOutRows: [] } }
      }]
    }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    const tables = back.sheets[0].layout?.tables
    expect(tables?.length).toBe(1)
    const rt = tables![0]
    expect(rt.name).toBe('Table1')
    expect(rt.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 })
    expect(rt.styleName).toBe('TableStyleMedium9')
    expect(rt.options.totalRow).toBe(true)
    expect(rt.options.bandedColumns).toBe(true)
    expect(rt.options.firstColumn).toBe(true)
    expect(rt.columns.map((c) => c.name)).toEqual(['Item', 'Qty', 'Total'])
    expect(rt.columns[1].totalsRowFunction).toBe('sum')
    expect(rt.columns[2].calculatedFormula).toBe('=B2*2')
    // Cell content itself round-trips too (tables don't touch how cells are written).
    expect(back.sheets[0].cells['1,0'].raw).toBe('A')
    expect(back.sheets[0].cells['3,1'].raw).toBe('=SUBTOTAL(109,B2:B3)')
  })

  it('round-trips a table with no Total Row and Filter Button off', async () => {
    const t = makeTable({ options: { ...defaultTableStyleOptions(), totalRow: false, filterButton: false } })
    const data: WorkbookData = {
      sheets: [{ name: 'S', cells: { '0,0': { raw: 'Item' } }, layout: { tables: [t] } }]
    }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    const rt = back.sheets[0].layout?.tables?.[0]
    expect(rt?.options.totalRow).toBe(false)
    expect(rt?.options.filterButton).toBe(false)
    expect(rt?.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 })
  })

  it('round-trips two tables on the same sheet', async () => {
    const t1 = makeTable({ id: 'a', name: 'Sales', range: { r1: 0, c1: 0, r2: 1, c2: 1 }, columns: [{ name: 'X' }, { name: 'Y' }], options: { ...defaultTableStyleOptions() } })
    const t2 = makeTable({ id: 'b', name: 'Costs', range: { r1: 0, c1: 4, r2: 1, c2: 5 }, columns: [{ name: 'P' }, { name: 'Q' }], options: { ...defaultTableStyleOptions() } })
    const data: WorkbookData = {
      sheets: [{ name: 'S', cells: { '0,0': { raw: 'X' }, '0,4': { raw: 'P' } }, layout: { tables: [t1, t2] } }]
    }
    const bytes = await writeXlsx(data)
    const back = await readXlsx(bytes)
    const names = (back.sheets[0].layout?.tables ?? []).map((t) => t.name).sort()
    expect(names).toEqual(['Costs', 'Sales'])
  })
})

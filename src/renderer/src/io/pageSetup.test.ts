// Page Setup / printing round-trips: pageSetup, headerFooter, print area/titles (the
// _xlnm.Print_Area/_xlnm.Print_Titles defined names) and manual row/column page breaks all need to
// survive a save, and the fast reader (fastXlsx.ts) must agree exactly with the exceljs reader -
// same pattern as fastXlsx.test.ts's differential test.
import { describe, it, expect } from 'vitest'
import { writeXlsx, readXlsxExceljs, readXlsx } from './xlsx'
import { defaultLayout } from '../model/defaults'
import { defaultPageSetup } from '../model/types'
import type { WorkbookData } from '../model/store'
import { computePrintPlan } from '../print/paginate'
import { renderPageHtml } from '../print/renderPage'
import { resolveHFCode } from '../print/headerFooter'
import { WorkbookStore } from '../model/store'

function fullFeaturedLayout() {
  const layout = defaultLayout()
  const ps = layout.pageSetup!
  ps.orientation = 'landscape'
  ps.paperSize = 'a4'
  ps.margins = { top: 1, bottom: 1, left: 0.5, right: 0.5, header: 0.4, footer: 0.4 }
  ps.fitToPage = true
  ps.fitToWidth = 2
  ps.fitToHeight = 0
  ps.centerHorizontally = true
  ps.centerVertically = true
  ps.printArea = [{ r1: 0, c1: 0, r2: 10, c2: 5 }]
  ps.printTitleRows = { r1: 0, r2: 1 }
  ps.printTitleCols = { c1: 0, c2: 0 }
  ps.gridlines = true
  ps.headings = true
  ps.blackAndWhite = true
  ps.draft = true
  ps.pageOrder = 'overThenDown'
  ps.useFirstPageNumber = true
  ps.firstPageNumber = 3
  ps.headerFooter.oddHeader = { left: 'L', center: '&P of &N', right: 'R' }
  ps.headerFooter.oddFooter = { left: '', center: '&F', right: '' }
  ps.headerFooter.evenHeader = { left: 'EL', center: '', right: '' }
  ps.headerFooter.firstHeader = { left: '', center: 'FIRST', right: '' }
  ps.headerFooter.differentOddEven = true
  ps.headerFooter.differentFirst = true
  ps.rowBreaks = [5, 10]
  ps.colBreaks = [3]
  return layout
}

describe('PageSetup .xlsx round-trip', () => {
  it('default pageSetup survives a save (both readers)', async () => {
    const data: WorkbookData = { sheets: [{ name: 'Sheet1', cells: { '0,0': { raw: 'hi' } }, layout: defaultLayout() }], activeSheet: 0 }
    const bytes = await writeXlsx(data)
    const slow = await readXlsxExceljs(bytes)
    const fast = await readXlsx(bytes)
    expect(slow.sheets[0].layout?.pageSetup).toEqual(defaultPageSetup())
    expect(fast.sheets[0].layout?.pageSetup).toEqual(defaultPageSetup())
  })

  it('a fully-customized pageSetup (margins, paper, scale, print area/titles, gridlines/headings, B&W, draft, page order, first page number, header/footer incl. odd/even/first, row+col breaks) round-trips exactly, and the fast reader agrees with the exceljs reader', async () => {
    const layout = fullFeaturedLayout()
    const data: WorkbookData = { sheets: [{ name: 'Sheet1', cells: { '0,0': { raw: 'hi' } }, layout }], activeSheet: 0 }
    const bytes = await writeXlsx(data)
    const slow = await readXlsxExceljs(bytes)
    const fast = await readXlsx(bytes)

    expect(fast.sheets[0].layout?.pageSetup).toEqual(slow.sheets[0].layout?.pageSetup)
    expect(slow.sheets[0].layout?.pageSetup).toEqual(layout.pageSetup)
  })

  it('print area with multiple disjoint ranges (Add to Print Area) round-trips', async () => {
    const layout = defaultLayout()
    layout.pageSetup!.printArea = [{ r1: 0, c1: 0, r2: 2, c2: 2 }, { r1: 5, c1: 5, r2: 8, c2: 8 }]
    const data: WorkbookData = { sheets: [{ name: 'Sheet1', cells: {}, layout }], activeSheet: 0 }
    const bytes = await writeXlsx(data)
    const slow = await readXlsxExceljs(bytes)
    expect(slow.sheets[0].layout?.pageSetup?.printArea).toEqual(layout.pageSetup!.printArea)
  })

  it('multi-sheet workbook keeps each sheet\'s own pageSetup/breaks distinct', async () => {
    const l1 = defaultLayout()
    l1.pageSetup!.orientation = 'landscape'
    l1.pageSetup!.rowBreaks = [3]
    const l2 = defaultLayout()
    l2.pageSetup!.orientation = 'portrait'
    l2.pageSetup!.colBreaks = [2]
    const data: WorkbookData = {
      sheets: [{ name: 'A', cells: {}, layout: l1 }, { name: 'B', cells: {}, layout: l2 }],
      activeSheet: 0
    }
    const bytes = await writeXlsx(data)
    const fast = await readXlsx(bytes)
    expect(fast.sheets[0].layout?.pageSetup?.orientation).toBe('landscape')
    expect(fast.sheets[0].layout?.pageSetup?.rowBreaks).toEqual([3])
    expect(fast.sheets[1].layout?.pageSetup?.orientation).toBe('portrait')
    expect(fast.sheets[1].layout?.pageSetup?.colBreaks).toEqual([2])
  })
})

describe('header/footer field codes', () => {
  it('resolves &P &N &D &T &F &A', () => {
    const ctx = { page: 2, pageCount: 5, date: new Date(2024, 0, 15), fileName: 'Book1.xlsx', sheetName: 'Sheet1', filePath: 'Book1.xlsx' }
    expect(resolveHFCode('Page &P of &N', ctx)).toBe('Page 2 of 5')
    expect(resolveHFCode('&A - &F', ctx)).toBe('Sheet1 - Book1.xlsx')
    expect(resolveHFCode('&D', ctx)).toContain('2024')
  })
})

describe('pagination engine', () => {
  function makeStore() {
    const store = new WorkbookStore()
    return { store, sheet: store.active }
  }

  it('a small used range fits on one page', () => {
    const { store, sheet } = makeStore()
    store.setRaw(sheet, 0, 0, 'hi')
    const plan = computePrintPlan(store, sheet)
    expect(plan.empty).toBe(false)
    expect(plan.pages.length).toBe(1)
    expect(plan.pages[0].pageNumber).toBe(1)
  })

  it('paginates a wide range into multiple pages across the printable width', () => {
    const { store, sheet } = makeStore()
    for (let c = 0; c < 40; c++) store.setRaw(sheet, 0, c, 'x')
    const plan = computePrintPlan(store, sheet)
    expect(plan.pages.length).toBeGreaterThan(1)
  })

  it('fit-to-1-page-wide scales content down to a single column of pages', () => {
    const { store, sheet } = makeStore()
    for (let c = 0; c < 40; c++) store.setRaw(sheet, 0, c, 'x')
    store.updateLayout(sheet, 'Scale to Fit', (l) => { l.pageSetup!.fitToPage = true; l.pageSetup!.fitToWidth = 1; l.pageSetup!.fitToHeight = 0 })
    const plan = computePrintPlan(store, sheet)
    expect(plan.scalePercent).toBeLessThan(100)
    expect(new Set(plan.pages.map((p) => p.colStart)).size).toBe(1)
  })

  it('respects a manual row break', () => {
    const { store, sheet } = makeStore()
    store.setRaw(sheet, 0, 0, 'a')
    store.setRaw(sheet, 5, 0, 'b')
    store.updateLayout(sheet, 'Insert Page Break', (l) => { l.pageSetup!.rowBreaks = [3] })
    const plan = computePrintPlan(store, sheet)
    expect(plan.pages[0].rowEnd).toBe(2)
    expect(plan.pages[1].rowStart).toBe(3)
  })

  it('an empty sheet with no print area produces no pages', () => {
    const { store, sheet } = makeStore()
    const plan = computePrintPlan(store, sheet)
    expect(plan.empty).toBe(true)
    expect(plan.pages.length).toBe(0)
  })

  it('renders a page to HTML with the resolved header/footer and cell text', () => {
    const { store, sheet } = makeStore()
    store.setRaw(sheet, 0, 0, 'Hello')
    store.updateLayout(sheet, 'Header', (l) => { l.pageSetup!.headerFooter.oddHeader = { left: '', center: '&A', right: '' } })
    const plan = computePrintPlan(store, sheet)
    const html = renderPageHtml(store, sheet, sheet.layout.pageSetup!, {
      page: plan.pages[0],
      pageCount: plan.pages.length,
      scalePercent: plan.scalePercent,
      hfCtx: { date: new Date(), fileName: 'Book1.xlsx', sheetName: sheet.name, filePath: 'Book1.xlsx' }
    })
    expect(html).toContain('Hello')
    expect(html).toContain(sheet.name)
  })
})

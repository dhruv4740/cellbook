import { describe, expect, it } from 'vitest'
import { paginate, PaginateInput } from './paginate'
import { buildPageSetupXml, parsePageSetupXml } from './pageSetupXml'
import { hfRuns, splitHf, joinHf, marginPresetOf, MARGIN_PRESETS } from '../model/pageSetup'
import { hfFromDisplay, hfToDisplay } from '../dialogs/PageSetupDialog'
import { WorkbookStore } from '../model/store'
import { readXlsx, writeXlsx } from './xlsx'
import { PageSetup } from '../model/types'

// 96 dpi: 20px rows (0.208in), 64px columns (0.667in)
const base = (o: Partial<PaginateInput> = {}): PaginateInput => ({
  range: { r1: 0, c1: 0, r2: 99, c2: 9 },
  colPx: () => 96, rowPx: () => 24, areaW: 6.5, areaH: 9, scale: 100,
  rowBreaks: [], colBreaks: [], order: 'downThenOver', ...o
})

describe('paginate', () => {
  it('fills pages greedily by height and width', () => {
    const p = paginate(base())
    expect(p.pagesTall).toBe(Math.ceil(100 / 36)) // 9in / 0.25in = 36 rows
    expect(p.pagesWide).toBe(2) // 6.5in / 1in = 6 columns
    expect(p.pages.length).toBe(p.pagesTall * 2)
  })
  it('orders down-then-over vs over-then-down', () => {
    const d = paginate(base()).pages, o = paginate(base({ order: 'overThenDown' })).pages
    expect(d[1].c1).toBe(d[0].c1)
    expect(o[1].c1).toBeGreaterThan(o[0].c1)
  })
  it('honours manual breaks', () => {
    const p = paginate(base({ rowBreaks: [10], colBreaks: [3] }))
    expect(p.rowSpans[0]).toEqual([0, 9])
    expect(p.colSpans[0]).toEqual([0, 2])
  })
  it('fits to one page wide, automatic tall, and ignores manual breaks', () => {
    const p = paginate(base({ fit: { w: 1, h: 0 }, rowBreaks: [5] }))
    expect(p.pagesWide).toBe(1)
    expect(p.scale).toBeLessThan(100)
    expect(p.rowSpans[0][1]).toBeGreaterThan(5)
  })
  it('repeats title rows only on later pages', () => {
    const p = paginate(base({ titleRows: [0, 0] }))
    expect(p.pages[0].repeatRows).toBe(false)
    expect(p.pages[1].repeatRows).toBe(true)
    expect(p.rowSpans[1][1] - p.rowSpans[1][0] + 1).toBe(35) // one row is spent on the title
  })
})

describe('header and footer codes', () => {
  it('splits and joins sections', () => {
    const p = splitHf('&LLeft&CMid&RPage &P')
    expect(p).toEqual({ left: 'Left', center: 'Mid', right: 'Page &P' })
    expect(joinHf(p)).toBe('&LLeft&CMid&RPage &P')
    expect(splitHf('Plain').center).toBe('Plain')
  })
  it('expands page fields and fonts', () => {
    const ctx = { page: 2, pages: 5, sheet: 'S', file: 'F.xlsx', path: 'C:/', now: new Date(2026, 0, 2) }
    expect(hfRuns('Page &P of &N', ctx).map((r) => r.text).join('')).toBe('Page 2 of 5')
    const r = hfRuns('&"Arial,Bold"&14Hi &&', ctx)
    expect(r[0]).toMatchObject({ text: 'Hi &', bold: true, size: 14, font: 'Arial' })
  })
  it('converts between file codes and the editor form', () => {
    expect(hfToDisplay('Page &P of &N')).toBe('Page &[Page] of &[Pages]')
    expect(hfFromDisplay('Page &[Page] of &[Pages] &&')).toBe('Page &P of &N &&')
  })
})

const sample: PageSetup = {
  landscape: true, paperSize: 8, margins: { left: 0.5, right: 0.6, top: 1, bottom: 0.9, header: 0.4, footer: 0.2 },
  scale: 80, gridlines: true, headings: true, centerH: true, blackAndWhite: true, pageOrder: 'overThenDown', firstPageNumber: 3,
  comments: 'atEnd', errors: 'dash',
  hf: { oddHeader: '&CTitle', oddFooter: '&LL&RPage &P', differentFirst: true, firstHeader: '&Cfirst', scaleWithDoc: false, alignWithMargins: false },
  rowBreaks: [10, 20], colBreaks: [4]
}

describe('page setup xml', () => {
  it('round-trips through the sheet xml', () => {
    const xml = `<worksheet><sheetData/>${buildPageSetupXml(sample)}</worksheet>`
    expect(parsePageSetupXml(xml)).toMatchObject(sample)
  })
  it('round-trips fit-to-page', () => {
    const ps: PageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    const xml = `<worksheet><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/>${buildPageSetupXml(ps)}</worksheet>`
    expect(parsePageSetupXml(xml)).toMatchObject(ps)
  })
  it('recognises margin presets', () => {
    expect(marginPresetOf({ margins: MARGIN_PRESETS.wide })).toBe('wide')
    expect(marginPresetOf({ margins: { left: 0.1 } })).toBe('custom')
    expect(marginPresetOf(undefined)).toBe('normal')
  })
  it('survives a real xlsx save, with print titles', async () => {
    const s = new WorkbookStore()
    const sh = s.active
    s.setRaw(sh, 0, 0, 'x')
    s.updateLayout(sh, 'Page Setup', (l) => { l.pageSetup = sample })
    s.setPrintTitles(sh, { rows: [0, 1], cols: [0, 0] })
    const back = await readXlsx(await writeXlsx(s.toData()))
    expect(back.sheets[0].layout?.pageSetup).toMatchObject(sample)
    const t = back.names?.find((n) => /Print_Titles/.test(n.name))
    expect(t?.ref).toMatch(/\$A:\$A,.*\$1:\$2/)
  })
})

// Shared PageSetup <-> .xlsx conversions, used by both the exceljs writer/reader (io/xlsx.ts) and
// the fast fflate-based reader (io/fastXlsx.ts / fastXlsxXml.ts), so the two stay in agreement.
// Print area (_xlnm.Print_Area) and print titles (_xlnm.Print_Titles) are workbook-scoped defined
// names in OOXML - exceljs round-trips them through ws.pageSetup.printArea/printTitlesRow/Column
// automatically; the fast reader parses xl/workbook.xml's <definedNames> directly (see fastXlsx.ts).
import {
  PageSetup, PaperSize, PAPER_SIZES, HeaderFooterSetup, HeaderFooterText, emptyHFText,
  defaultPageSetup, defaultHeaderFooter, colName, colIndex, rangeA1, Range
} from '../model/types'

const OOXML_TO_PAPER: Record<number, PaperSize> = {}
for (const [k, v] of Object.entries(PAPER_SIZES)) OOXML_TO_PAPER[v.ooxml] = k as PaperSize

export function paperSizeToOoxml(p: PaperSize): number {
  return PAPER_SIZES[p]?.ooxml ?? 1
}
export function ooxmlToPaperSize(n: number | undefined): PaperSize {
  return (n !== undefined && OOXML_TO_PAPER[n]) || 'letter'
}

// ---------- header/footer: Excel stores one string per odd/even/first header/footer, with
// &L / &C / &R markers introducing each section (any section can be omitted). ----------
export function combineHF(t: HeaderFooterText): string | undefined {
  let s = ''
  if (t.left) s += '&L' + t.left
  if (t.center) s += '&C' + t.center
  if (t.right) s += '&R' + t.right
  return s || undefined
}

export function splitHF(s: string | undefined | null): HeaderFooterText {
  const out = emptyHFText()
  if (!s) return out
  const parts = s.split(/(&[LCR])/)
  let cur: 'left' | 'center' | 'right' | null = null
  for (const part of parts) {
    if (part === '&L') cur = 'left'
    else if (part === '&C') cur = 'center'
    else if (part === '&R') cur = 'right'
    else if (cur) out[cur] += part
  }
  return out
}

// ---------- PageSetup <-> exceljs's Partial<PageSetup>/Partial<HeaderFooter> ----------
export interface ExceljsPageSetupLike {
  orientation?: 'portrait' | 'landscape'
  fitToPage?: boolean
  fitToWidth?: number
  fitToHeight?: number
  scale?: number
  paperSize?: number
  margins?: { top: number; bottom: number; left: number; right: number; header: number; footer: number }
  horizontalCentered?: boolean
  verticalCentered?: boolean
  showRowColHeaders?: boolean
  showGridLines?: boolean
  blackAndWhite?: boolean
  draft?: boolean
  pageOrder?: 'downThenOver' | 'overThenDown'
  firstPageNumber?: number
  /** Not in exceljs's published TS types, but its own xform reads/writes it (page-setup-xform.js) -
   *  firstPageNumber itself always comes back as 1 when absent, so this is the only reliable signal. */
  useFirstPageNumber?: boolean
  printArea?: string
  printTitlesRow?: string
  printTitlesColumn?: string
}
export interface ExceljsHeaderFooterLike {
  differentFirst?: boolean
  differentOddEven?: boolean
  oddHeader?: string | null
  oddFooter?: string | null
  evenHeader?: string | null
  evenFooter?: string | null
  firstHeader?: string | null
  firstFooter?: string | null
}

export function toExceljsPageSetup(ps: PageSetup): ExceljsPageSetupLike {
  return {
    orientation: ps.orientation,
    fitToPage: ps.fitToPage,
    // 0 is a meaningful value here ("Automatic" - unconstrained on that axis), so no `|| 1` fallback.
    fitToWidth: ps.fitToWidth,
    fitToHeight: ps.fitToHeight,
    scale: ps.scalePercent,
    paperSize: paperSizeToOoxml(ps.paperSize),
    margins: { ...ps.margins },
    horizontalCentered: ps.centerHorizontally,
    verticalCentered: ps.centerVertically,
    showRowColHeaders: ps.headings,
    showGridLines: ps.gridlines,
    blackAndWhite: ps.blackAndWhite,
    draft: ps.draft,
    pageOrder: ps.pageOrder,
    firstPageNumber: ps.useFirstPageNumber ? ps.firstPageNumber : undefined,
    // exceljs concatenates multiple disjoint print areas with '&&' (see workbook-xform.js), writing
    // one _xlnm.Print_Area defined name per range and joining them the same way on read.
    printArea: ps.printArea.length ? ps.printArea.map(rangeA1).join('&&') : undefined,
    printTitlesRow: ps.printTitleRows ? `${ps.printTitleRows.r1 + 1}:${ps.printTitleRows.r2 + 1}` : undefined,
    printTitlesColumn: ps.printTitleCols ? `${colName(ps.printTitleCols.c1)}:${colName(ps.printTitleCols.c2)}` : undefined
  }
}

export function toExceljsHeaderFooter(hf: HeaderFooterSetup): ExceljsHeaderFooterLike {
  return {
    differentFirst: hf.differentFirst,
    differentOddEven: hf.differentOddEven,
    oddHeader: combineHF(hf.oddHeader),
    oddFooter: combineHF(hf.oddFooter),
    evenHeader: combineHF(hf.evenHeader),
    evenFooter: combineHF(hf.evenFooter),
    firstHeader: combineHF(hf.firstHeader),
    firstFooter: combineHF(hf.firstFooter)
  }
}

function parseRangeList(s: string | undefined): Range[] {
  if (!s) return []
  return s.split('&&').map((part) => {
    const clean = part.replace(/\$/g, '').replace(/^[^!]*!/, '')
    const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(clean.trim())
    if (!m) return null
    return { r1: +m[2] - 1, c1: colIndex(m[1]), r2: +m[4] - 1, c2: colIndex(m[3]) }
  }).filter((r): r is Range => r !== null)
}

/** Parses a defined-name-style print titles value, e.g. `Sheet1!$1:$3` or `1:3`. */
export function parseTitleRows(s: string | undefined): { r1: number; r2: number } | undefined {
  if (!s) return undefined
  const m = /(\d+):(\d+)/.exec(s.replace(/\$/g, ''))
  return m ? { r1: +m[1] - 1, r2: +m[2] - 1 } : undefined
}
export function parseTitleCols(s: string | undefined): { c1: number; c2: number } | undefined {
  if (!s) return undefined
  const m = /([A-Za-z]+):([A-Za-z]+)/.exec(s.replace(/\$/g, ''))
  return m ? { c1: colIndex(m[1]), c2: colIndex(m[2]) } : undefined
}

export function fromExceljsPageSetup(exps: ExceljsPageSetupLike | undefined, hf: ExceljsHeaderFooterLike | undefined, rowBreaks: number[], colBreaks: number[]): PageSetup {
  const base = defaultPageSetup()
  if (!exps) return base
  return {
    orientation: exps.orientation ?? base.orientation,
    paperSize: ooxmlToPaperSize(exps.paperSize),
    fitToPage: exps.fitToPage ?? base.fitToPage,
    scalePercent: exps.scale ?? base.scalePercent,
    fitToWidth: exps.fitToWidth ?? base.fitToWidth,
    fitToHeight: exps.fitToHeight ?? base.fitToHeight,
    margins: exps.margins ? { ...exps.margins } : base.margins,
    centerHorizontally: exps.horizontalCentered ?? base.centerHorizontally,
    centerVertically: exps.verticalCentered ?? base.centerVertically,
    printArea: parseRangeList(exps.printArea),
    printTitleRows: parseTitleRows(exps.printTitlesRow),
    printTitleCols: parseTitleCols(exps.printTitlesColumn),
    gridlines: exps.showGridLines ?? base.gridlines,
    headings: exps.showRowColHeaders ?? base.headings,
    blackAndWhite: exps.blackAndWhite ?? base.blackAndWhite,
    draft: exps.draft ?? base.draft,
    pageOrder: exps.pageOrder ?? base.pageOrder,
    useFirstPageNumber: !!exps.useFirstPageNumber,
    firstPageNumber: exps.useFirstPageNumber ? (exps.firstPageNumber ?? base.firstPageNumber) : base.firstPageNumber,
    headerFooter: fromExceljsHeaderFooter(hf),
    rowBreaks,
    colBreaks
  }
}

export function fromExceljsHeaderFooter(hf: ExceljsHeaderFooterLike | undefined): HeaderFooterSetup {
  const base = defaultHeaderFooter()
  if (!hf) return base
  return {
    ...base,
    differentFirst: !!hf.differentFirst,
    differentOddEven: !!hf.differentOddEven,
    oddHeader: splitHF(hf.oddHeader),
    oddFooter: splitHF(hf.oddFooter),
    evenHeader: splitHF(hf.evenHeader),
    evenFooter: splitHF(hf.evenFooter),
    firstHeader: splitHF(hf.firstHeader),
    firstFooter: splitHF(hf.firstFooter)
  }
}

/** exceljs's RowBreak model shape ({id, min, max, man}); `id` is the 1-based row *after* which the
 *  break falls, matching our 0-based rowBreaks (row index that starts a new page). */
export function rowBreaksToExceljs(rowBreaks: number[]): { id: number; min: number; max: number; man: number }[] {
  return [...rowBreaks].sort((a, b) => a - b).map((r) => ({ id: r, min: 0, max: 16383, man: 1 }))
}
export function exceljsRowBreaksToArray(model: { id: number }[] | undefined): number[] {
  return (model ?? []).map((b) => b.id)
}

// Renders one PrintPage to an HTML string: a table sized/positioned exactly like Excel's printed
// page (repeated title rows/cols, cell formatting - fonts, fills, borders, number formats, merges,
// wrap, alignment - plus header/footer text). Used by the Print backstage preview, the hidden print
// BrowserWindow, printToPDF export and the browser window.print() fallback alike, so what you see
// in preview is exactly what prints.
import { CellStyle, PageSetup, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, DEFAULT_FONT, DEFAULT_FONT_SIZE } from '../model/types'
import { WorkbookStore, SheetState } from '../model/store'
import { PrintPage } from './paginate'
import { pageGeometry } from './pageGeometry'
import { resolveHFText, HFContext } from './headerFooter'

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function borderCss(side: { style: string; color?: string } | undefined): string {
  if (!side) return ''
  const width = side.style === 'thick' || side.style === 'mediumDashDot' || side.style === 'mediumDashDotDot' || side.style.startsWith('medium') ? '2px' : side.style === 'hair' ? '0.5px' : '1px'
  const dash = side.style.includes('Dash') || side.style === 'dashed' || side.style === 'dotted' ? (side.style === 'dotted' ? 'dotted' : 'dashed') : side.style === 'double' ? 'double' : 'solid'
  return `${side.style === 'double' ? '3px' : width} ${dash} ${side.color ?? '#000'}`
}

function cellCss(style: CellStyle | undefined, blackAndWhite: boolean): string {
  const css: string[] = []
  css.push(`font-family:${(style?.fontFamily ?? DEFAULT_FONT)},sans-serif`)
  css.push(`font-size:${(style?.fontSize ?? DEFAULT_FONT_SIZE)}pt`)
  if (style?.bold) css.push('font-weight:bold')
  if (style?.italic) css.push('font-style:italic')
  if (style?.underline) css.push(`text-decoration:${style.underline === 'double' ? 'underline' : 'underline'}`)
  if (style?.strike) css.push('text-decoration:line-through')
  if (!blackAndWhite && style?.color) css.push(`color:${style.color}`)
  if (!blackAndWhite && style?.fill) css.push(`background-color:${style.fill}`)
  const ha = style?.hAlign && style.hAlign !== 'general' ? style.hAlign : undefined
  if (ha) css.push(`text-align:${ha === 'centerContinuous' ? 'center' : ha}`)
  const va = style?.vAlign ?? 'bottom'
  css.push(`vertical-align:${va === 'middle' ? 'middle' : va}`)
  if (style?.wrap) css.push('white-space:pre-wrap;word-break:break-word')
  else css.push('white-space:pre;overflow:hidden')
  if (style?.borders) {
    const b = style.borders
    if (b.top) css.push(`border-top:${borderCss(b.top)}`)
    if (b.right) css.push(`border-right:${borderCss(b.right)}`)
    if (b.bottom) css.push(`border-bottom:${borderCss(b.bottom)}`)
    if (b.left) css.push(`border-left:${borderCss(b.left)}`)
  }
  return css.join(';')
}

/** Renders one axis (all title cols + the page's content cols, or title rows + content rows) as an
 *  ordered list of indices, deduped so a title range overlapping the content range isn't repeated. */
function axisIndices(titleRange: { a: number; b: number } | undefined, start: number, end: number): number[] {
  const out: number[] = []
  const seen = new Set<number>()
  if (titleRange) for (let i = titleRange.a; i <= titleRange.b; i++) if (!seen.has(i)) { out.push(i); seen.add(i) }
  for (let i = start; i <= end; i++) if (!seen.has(i)) { out.push(i); seen.add(i) }
  return out
}

export interface RenderPageOptions {
  page: PrintPage
  pageCount: number
  hfCtx: Omit<HFContext, 'page' | 'pageCount'>
  /** The plan's effective scale percent (resolveScale in paginate.ts): 100 = actual size. */
  scalePercent: number
}

export function renderPageHtml(store: WorkbookStore, sheet: SheetState, ps: PageSetup, opts: RenderPageOptions): string {
  const geo = pageGeometry(ps)
  const factor = (opts.scalePercent || 100) / 100
  const rows = axisIndices(ps.printTitleRows ? { a: ps.printTitleRows.r1, b: ps.printTitleRows.r2 } : undefined, opts.page.rowStart, opts.page.rowEnd)
  const cols = axisIndices(ps.printTitleCols ? { a: ps.printTitleCols.c1, b: ps.printTitleCols.c2 } : undefined, opts.page.colStart, opts.page.colEnd)
  const colIndexOf = new Map(cols.map((c, i) => [c, i]))
  const rowIndexOf = new Map(rows.map((r, i) => [r, i]))

  const colWidths = cols.map((c) => (sheet.layout.hiddenCols[c] ? 0 : (sheet.layout.colWidths[c] ?? DEFAULT_COL_WIDTH)))
  const rowHeights = rows.map((r) => (sheet.layout.hiddenRows[r] ? 0 : (sheet.layout.rowHeights[r] ?? DEFAULT_ROW_HEIGHT)))
  const totalW = colWidths.reduce((a, b) => a + b, 0) * factor
  const totalH = rowHeights.reduce((a, b) => a + b, 0) * factor

  // rendered merges: any merge whose top-left cell is within this page's row/col set.
  const merges = sheet.layout.merges.filter((m) => rowIndexOf.has(m.r1) && colIndexOf.has(m.c1))
  const covered = new Set<string>()
  for (const m of merges) {
    for (let r = m.r1; r <= m.r2; r++)
      for (let c = m.c1; c <= m.c2; c++)
        if (!(r === m.r1 && c === m.c1)) covered.add(r + ',' + c)
  }
  const mergeAt = new Map(merges.map((m) => [m.r1 + ',' + m.c1, m]))

  let rowsHtml = ''
  for (const r of rows) {
    const h = sheet.layout.hiddenRows[r] ? 0 : (sheet.layout.rowHeights[r] ?? DEFAULT_ROW_HEIGHT)
    if (h === 0) continue
    let rowCells = ''
    for (const c of cols) {
      if (sheet.layout.hiddenCols[c]) continue
      if (covered.has(r + ',' + c)) continue
      const m = mergeAt.get(r + ',' + c)
      let colspan = 1, rowspan = 1
      if (m) {
        for (let cc = m.c1; cc <= Math.min(m.c2, cols[cols.length - 1]); cc++) if (colIndexOf.has(cc)) colspan = colIndexOf.get(cc)! - colIndexOf.get(m.c1)! + 1
        for (let rr = m.r1; rr <= Math.min(m.r2, rows[rows.length - 1]); rr++) if (rowIndexOf.has(rr)) rowspan = rowIndexOf.get(rr)! - rowIndexOf.get(m.r1)! + 1
      }
      const style = store.getStyle(sheet, r, c)
      const disp = store.getDisplay(sheet, r, c)
      const css = cellCss(style, ps.blackAndWhite)
      const attrs = (colspan > 1 ? ` colspan="${colspan}"` : '') + (rowspan > 1 ? ` rowspan="${rowspan}"` : '')
      rowCells += `<td style="${css}"${attrs}>${esc(disp.text)}</td>`
    }
    rowsHtml += `<tr style="height:${h}px">${rowCells}</tr>`
  }

  const colgroup = colWidths.map((w) => `<col style="width:${w}px">`).join('')

  const showGridlines = ps.gridlines
  const showHeadings = ps.headings
  let headerRowHtml = ''
  if (showHeadings) {
    headerRowHtml = '<tr><td class="xl-print-corner"></td>' + cols.map((c) => `<td class="xl-print-colhead">${colLetters(c)}</td>`).join('') + '</tr>'
  }
  let bodyWithHeadings = rowsHtml
  if (showHeadings) {
    const lines = rowsHtml.split('</tr>').filter(Boolean)
    bodyWithHeadings = lines.map((line, i) => line.replace('<tr', `<tr><td class="xl-print-rowhead">${rows[i] + 1}</td>`) + '</tr>').join('')
  }

  const ctx: HFContext = { ...opts.hfCtx, page: opts.page.pageNumber, pageCount: opts.pageCount }
  const hf = ps.headerFooter
  const firstPageNum = ps.useFirstPageNumber ? ps.firstPageNumber : 1
  const isFirstPage = opts.page.pageNumber === firstPageNum
  const isEvenPage = opts.page.pageNumber % 2 === 0
  const hSrc = hf.differentFirst && isFirstPage ? hf.firstHeader : hf.differentOddEven && isEvenPage ? hf.evenHeader : hf.oddHeader
  const fSrc = hf.differentFirst && isFirstPage ? hf.firstFooter : hf.differentOddEven && isEvenPage ? hf.evenFooter : hf.oddFooter
  const header = resolveHFText(hSrc, ctx)
  const footer = resolveHFText(fSrc, ctx)

  const marginLeftForCenter = ps.centerHorizontally ? Math.max(0, (geo.printableWidthPx - totalW) / 2) : 0
  const marginTopForCenter = ps.centerVertically ? Math.max(0, (geo.printableHeightPx - totalH) / 2) : 0

  const gridlineClass = showGridlines ? 'xl-print-gridlines' : ''

  return `
<div class="xl-print-page" style="width:${geo.paperWidthPx}px;height:${geo.paperHeightPx}px;padding:${geo.marginTopPx}px ${geo.marginRightPx}px ${geo.marginBottomPx}px ${geo.marginLeftPx}px;${ps.draft ? 'filter:grayscale(1) contrast(0.6);' : ''}">
  <div class="xl-print-header" style="height:${geo.headerPx}px">
    <div class="xl-hf-l">${esc(header.left)}</div><div class="xl-hf-c">${esc(header.center)}</div><div class="xl-hf-r">${esc(header.right)}</div>
  </div>
  <div class="xl-print-body" style="padding-left:${marginLeftForCenter}px;padding-top:${marginTopForCenter}px">
    <table class="xl-print-table ${gridlineClass}" style="border-collapse:collapse;transform-origin:top left;transform:scale(${factor})">
      <colgroup>${showHeadings ? '<col style="width:28px">' : ''}${colgroup}</colgroup>
      <tbody>${headerRowHtml}${bodyWithHeadings}</tbody>
    </table>
  </div>
  <div class="xl-print-footer" style="height:${geo.footerPx}px">
    <div class="xl-hf-l">${esc(footer.left)}</div><div class="xl-hf-c">${esc(footer.center)}</div><div class="xl-hf-r">${esc(footer.right)}</div>
  </div>
</div>`
}

function colLetters(c: number): string {
  let s = '', n = c + 1
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) }
  return s
}

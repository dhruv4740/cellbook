// Builds the document File > Print and Export > PDF send to the printer. Pages are laid out here, the way
// Excel does (page setup: paper, margins, scale / fit, titles, breaks, order, headers & footers), so the
// preview, the PDF and the paper all show the same pages.
import { store, SheetState } from '../model/store'
import { Range, normRange, CellStyle, BorderSide, colName, PageSetup } from '../model/types'
import { evaluateCf } from '../format/conditional'
import { DEFAULT_PAPER, HfContext, HfRun, hfRuns, marginsOf, pageInches, paperOf, splitHf } from '../model/pageSetup'
import { paginate, PrintPage } from './paginate'

export interface PrintOptions {
  what: 'active' | 'workbook' | 'selection'
}
export const DEFAULT_PRINT: PrintOptions = { what: 'active' }

export interface PrintDoc {
  html: string
  /** Paper size of the first printed sheet, in inches (after orientation). */
  widthIn: number
  heightIn: number
  pageCount: number
}

/** Excel prints columns a little wider than their screen width (measured against Excel's page counts: ~3%). */
const PRINT_COL_FACTOR = 1.03
const pcw = (sheet: SheetState, c: number): number => store.colWidth(sheet, c) * PRINT_COL_FACTOR

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
function border(b?: BorderSide): string | undefined {
  if (!b || b.style === ('none' as string)) return undefined
  const w = /thick/.test(b.style) ? '2.5px' : /medium/i.test(b.style) ? '1.5px' : '1px'
  const st = b.style === 'double' ? 'double' : /dash|dot/i.test(b.style) ? (/dot/i.test(b.style) ? 'dotted' : 'dashed') : 'solid'
  return `${b.style === 'double' ? '3px' : w} ${st} ${b.color ?? '#000'}`
}

/** The range to print for a sheet: the print area / selection / used range. */
export function printRange(sheet: SheetState, what: PrintOptions['what']): Range | null {
  if (what === 'selection') {
    const n = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
    return store.clampToUsed(sheet, n)
  }
  const area = store.names.find((n) => /^_xlnm\.Print_Area$/i.test(n.name) && n.sheetUid === sheet.uid)
  if (area) {
    const m = /\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(area.ref)
    if (m) {
      const col = (s: string) => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1
      return { r1: +m[2] - 1, c1: col(m[1]), r2: +m[4] - 1, c2: col(m[3]) }
    }
  }
  const used = store.usedRange(sheet)
  if (!used) return null
  // Include cells that only carry fills/borders, like Excel's print range.
  return { r1: 0, c1: 0, r2: used.r2, c2: used.c2 }
}

let measureCtx: CanvasRenderingContext2D | null | undefined
function textWidth(text: string, st: CellStyle | undefined): number {
  if (measureCtx === undefined) measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  if (!measureCtx) return text.length * 7
  measureCtx.font = `${st?.italic ? 'italic ' : ''}${st?.bold ? 'bold ' : ''}${Math.round(((st?.fontSize ?? 11) * 96) / 72)}px ${st?.fontFamily ?? 'Calibri'}, sans-serif`
  return measureCtx.measureText(text).width
}

function errorText(text: string, ps: PageSetup): string {
  if (!/^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)$/.test(text)) return text
  return ps.errors === 'blank' ? '' : ps.errors === 'dash' ? '--' : ps.errors === 'NA' ? '#N/A' : text
}

/** One page's cells as an HTML table: `rows`/`cols` are the visible line indexes to draw (titles first). */
function pageTable(sheet: SheetState, rows: number[], cols: number[], ps: PageSetup): { html: string; w: number; h: number } {
  const l = sheet.layout
  const merges = l.merges.map(normRange)
  const covered = new Set<string>()
  const span = new Map<string, { rs: number; cs: number }>()
  const rowSet = new Set(rows), colSet = new Set(cols)
  for (const m of merges) {
    if (m.r2 < rows[0] || m.r1 > rows[rows.length - 1] || m.c2 < cols[0] || m.c1 > cols[cols.length - 1]) continue
    const rs = rows.filter((r) => r >= m.r1 && r <= m.r2).length
    const cs = cols.filter((c) => c >= m.c1 && c <= m.c2).length
    if (rowSet.has(m.r1) && colSet.has(m.c1)) span.set(`${m.r1},${m.c1}`, { rs, cs })
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`)
  }
  const hd = !!ps.headings
  const HW = 40
  const w = cols.reduce((a, c) => a + pcw(sheet, c), 0) + (hd ? HW : 0)
  const h = rows.reduce((a, r) => a + store.rowHeight(sheet, r), 0) + (hd ? 20 : 0)
  const out: string[] = []
  out.push(`<table class="sheet${ps.gridlines && !ps.draft ? ' grid' : ''}"><colgroup>${hd ? `<col style="width:${HW}px">` : ''}${cols.map((c) => `<col style="width:${pcw(sheet, c)}px">`).join('')}</colgroup>`)
  if (hd) out.push(`<thead><tr style="height:20px"><th></th>${cols.map((c) => `<th>${colName(c)}</th>`).join('')}</tr></thead>`)
  out.push('<tbody>')
  for (const r of rows) {
    out.push(`<tr style="height:${store.rowHeight(sheet, r)}px">${hd ? `<th>${r + 1}</th>` : ''}`)
    for (let ci = 0; ci < cols.length; ci++) {
      const c = cols[ci]
      const k = `${r},${c}`
      if (covered.has(k)) continue
      const sp = span.get(k)
      const base: CellStyle | undefined = store.getStyle(sheet, r, c)
      const cf = evaluateCf(store, sheet, r, c)
      const st: CellStyle = { ...base, ...(cf.style ?? {}) }
      const d = store.getDisplay(sheet, r, c)
      const text = errorText(d.text, ps)
      const css: string[] = []
      const bg = cf.colorScale || st.fill
      if (bg) css.push(`background:${bg}`)
      if (st.bold) css.push('font-weight:bold')
      if (st.italic) css.push('font-style:italic')
      const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ')
      if (deco) css.push(`text-decoration:${deco}${st.underline === 'double' ? ' double' : ''}`)
      if (d.color || st.color) css.push(`color:${d.color ?? st.color}`)
      if (st.fontFamily) css.push(`font-family:'${st.fontFamily}',Calibri,sans-serif`)
      if (st.fontSize) css.push(`font-size:${st.fontSize}pt`)
      const h2 = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'centerContinuous' ? 'center' : st.hAlign === 'distributed' ? 'justify' : st.hAlign) : d.align ?? 'left'
      css.push(`text-align:${h2}`)
      css.push(`vertical-align:${st.vAlign === 'top' ? 'top' : st.vAlign === 'middle' ? 'middle' : 'bottom'}`)
      if (st.wrap) css.push('white-space:pre-wrap')
      if (st.indent) css.push(`padding-left:${3 + st.indent * 9}px`)
      const b = st.borders
      if (b) {
        const t = border(b.top), rr = border(b.right), bb = border(b.bottom), ll = border(b.left)
        if (t) css.push(`border-top:${t}`)
        if (rr) css.push(`border-right:${rr}`)
        if (bb) css.push(`border-bottom:${bb}`)
        if (ll) css.push(`border-left:${ll}`)
      }
      // Left-aligned text that doesn't fit runs on into the empty cells to its right, like Excel.
      let cs = sp && sp.cs > 1 ? sp.cs : 1
      if (!sp && text && !st.wrap && h2 === 'left' && typeof store.getValue(sheet, r, c) !== 'number') {
        let need = textWidth(text, st) + 6 - pcw(sheet, c)
        let j = ci + 1
        while (need > 0 && j < cols.length && !covered.has(`${r},${cols[j]}`) && store.getRaw(sheet, r, cols[j]) === '' && !l.merges.some((m) => { const n = normRange(m); return r >= n.r1 && r <= n.r2 && cols[j] >= n.c1 && cols[j] <= n.c2 })) {
          need -= pcw(sheet, cols[j]); cs++; j++
        }
        if (cs > 1) { for (let q = ci + 1; q < ci + cs; q++) covered.add(`${r},${cols[q]}`); css.push('overflow:visible') }
      }
      const attrs = (sp && sp.rs > 1 ? ` rowspan="${sp.rs}"` : '') + (cs > 1 ? ` colspan="${cs}"` : '')
      out.push(`<td${attrs}${css.length ? ` style="${css.join(';')}"` : ''}>${esc(text)}</td>`)
    }
    out.push('</tr>')
  }
  out.push('</tbody></table>')
  return { html: out.join(''), w, h }
}

function runsHtml(runs: HfRun[], sizeMul: number): string {
  return runs.map((r) => {
    const css: string[] = []
    if (r.bold) css.push('font-weight:bold')
    if (r.italic) css.push('font-style:italic')
    const deco = [r.underline ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ')
    if (deco) css.push(`text-decoration:${deco}${r.underline === 'double' ? ' double' : ''}`)
    if (r.font) css.push(`font-family:'${r.font}'`)
    if (r.size) css.push(`font-size:${(r.size * sizeMul).toFixed(2)}pt`)
    if (r.color) css.push(`color:${r.color}`)
    if (r.sup) css.push('vertical-align:super;font-size:0.7em')
    if (r.sub) css.push('vertical-align:sub;font-size:0.7em')
    return `<span${css.length ? ` style="${css.join(';')}"` : ''}>${esc(r.text).replace(/\n/g, '<br>')}</span>`
  }).join('')
}

/** Which header/footer string applies to a page (first page / even / odd). */
export function hfFor(ps: PageSetup, pageInSheet: number, pageNumber: number, kind: 'Header' | 'Footer'): string {
  const hf = ps.hf
  if (!hf) return ''
  const pick = (k: 'odd' | 'even' | 'first'): string | undefined => hf[`${k}${kind}` as keyof typeof hf] as string | undefined
  if (hf.differentFirst && pageInSheet === 0) return pick('first') ?? ''
  if (hf.differentOddEven && pageNumber % 2 === 0) return pick('even') ?? ''
  return pick('odd') ?? ''
}

interface SheetPages { sheet: SheetState; ps: PageSetup; pages: PrintPage[]; scale: number; rg: Range }

export function layoutSheet(sheet: SheetState, what: PrintOptions['what'], firstNumber?: number): SheetPages | null {
  const rg = printRange(sheet, what)
  if (!rg) return null
  const ps: PageSetup = sheet.layout.pageSetup ?? {}
  const m = marginsOf(ps)
  const pg = pageInches(ps)
  const titles = store.printTitles(sheet)
  const l = sheet.layout
  const p = paginate({
    range: rg,
    colPx: (c) => (l.hiddenCols[c] ? 0 : pcw(sheet, c)),
    rowPx: (r) => (l.hiddenRows[r] ? 0 : store.rowHeight(sheet, r)),
    areaW: Math.max(0.5, pg.w - m.left - m.right), areaH: Math.max(0.5, pg.h - m.top - m.bottom),
    scale: ps.scale ?? 100,
    fit: ps.fitToPage ? { w: ps.fitToWidth ?? 1, h: ps.fitToHeight ?? 1 } : undefined,
    titleRows: titles.rows, titleCols: titles.cols,
    rowBreaks: ps.rowBreaks ?? [], colBreaks: ps.colBreaks ?? [],
    order: ps.pageOrder === 'overThenDown' ? 'overThenDown' : 'downThenOver',
    firstPageNumber: firstNumber ?? ps.firstPageNumber
  })
  return { sheet, ps, pages: p.pages, scale: p.scale, rg }
}

/** A complete HTML document for printing / PDF (or, with `preview`, a scrollable stack of pages). */
export function buildPrint(o: PrintOptions, preview = false): PrintDoc | null {
  const sheets = o.what === 'workbook' ? store.sheets.filter((s) => !s.layout.sheetHidden) : [store.active]
  const laid: SheetPages[] = []
  let next = 1
  for (const sh of sheets) {
    const sp = layoutSheet(sh, o.what, sh.layout.pageSetup?.firstPageNumber ?? next)
    if (!sp || !sp.pages.length) continue
    laid.push(sp)
    next += sp.pages.length
  }
  if (!laid.length) return null
  const total = next - 1
  const file = (store.fileName || 'Book1')
  const now = new Date()
  const css: string[] = []
  const body: string[] = []
  let first: { w: number; h: number } | undefined
  laid.forEach((sp, si) => {
    const { sheet, ps } = sp
    const pg = pageInches(ps)
    first ??= { w: pg.w, h: pg.h }
    const m = marginsOf(ps)
    const l = sheet.layout
    const titles = store.printTitles(sheet)
    css.push(`@page s${si} { size: ${pg.w}in ${pg.h}in; margin: 0 }`)
    const rowsOf = (a: number, b: number): number[] => { const out: number[] = []; for (let r = a; r <= b; r++) if (!l.hiddenRows[r]) out.push(r); return out }
    const colsOf = (a: number, b: number): number[] => { const out: number[] = []; for (let c = a; c <= b; c++) if (!l.hiddenCols[c]) out.push(c); return out }
    const s = sp.scale / 100
    const areaW = (pg.w - m.left - m.right) * 96, areaH = (pg.h - m.top - m.bottom) * 96
    sp.pages.forEach((p, pi) => {
      const rows = [...(p.repeatRows && titles.rows ? rowsOf(titles.rows[0], titles.rows[1]) : []), ...rowsOf(p.r1, p.r2)]
      const cols = [...(p.repeatCols && titles.cols ? colsOf(titles.cols[0], titles.cols[1]) : []), ...colsOf(p.c1, p.c2)]
      if (!rows.length || !cols.length) return
      const t = pageTable(sheet, rows, cols, ps)
      const cw = t.w * s, ch = t.h * s
      const offX = ps.centerH ? Math.max(0, (areaW - cw) / 2) : 0
      const offY = ps.centerV ? Math.max(0, (areaH - ch) / 2) : 0
      const hfCtx: HfContext = { page: p.number, pages: total, sheet: sheet.name, file, path: store.filePath ?? '', now }
      const mul = ps.hf?.scaleWithDoc === false ? 1 : s
      const align = ps.hf?.alignWithMargins === false ? { l: 0, r: 0 } : { l: m.left, r: m.right }
      const hfBlock = (kind: 'Header' | 'Footer'): string => {
        const code = hfFor(ps, pi, p.number, kind)
        if (!code) return ''
        const parts = splitHf(code)
        const cell = (side: 'left' | 'center' | 'right'): string => parts[side] ? `<div class="hf-${side}">${runsHtml(hfRuns(parts[side], hfCtx), mul)}</div>` : ''
        const pos = kind === 'Header' ? `top:${m.header}in` : `bottom:${m.footer}in`
        return `<div class="hf" style="${pos};left:${align.l}in;right:${align.r}in;font-size:${(11 * mul).toFixed(2)}pt">${cell('left')}${cell('center')}${cell('right')}</div>`
      }
      body.push(`<div class="page s${si}${ps.blackAndWhite ? ' bw' : ''}" style="page:s${si};width:${pg.w}in;height:${pg.h}in">`
        + hfBlock('Header')
        + `<div class="content" style="left:${m.left}in;top:${m.top}in;width:${pg.w - m.left - m.right}in;height:${pg.h - m.top - m.bottom}in">`
        + `<div style="position:absolute;left:${offX}px;top:${offY}px;width:${t.w}px;transform:scale(${s});transform-origin:0 0">${t.html}</div></div>`
        + hfBlock('Footer') + '</div>')
    })
  })
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>
${css.join('\n')}
html, body { margin: 0; padding: 0; }
body { font: 11pt Calibri, 'Segoe UI', Arial, sans-serif; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; ${preview ? 'background:#e6e6e6;padding:8px 0;' : ''} }
.page { position: relative; overflow: hidden; background: #fff; ${preview ? 'margin:0 auto 12px;box-shadow:0 1px 4px rgba(0,0,0,.35);' : 'break-after: page;'} }
.page:last-child { break-after: auto; }
.content { position: absolute; overflow: hidden; }
.hf { position: absolute; display: flex; justify-content: space-between; align-items: flex-end; white-space: pre; line-height: 1.2; }
.hf-left { flex: 1; text-align: left; } .hf-center { flex: 1; text-align: center; } .hf-right { flex: 1; text-align: right; }
table.sheet { border-collapse: collapse; table-layout: fixed; }
table.sheet td { padding: 0 3px; overflow: hidden; white-space: pre; line-height: 1.15; }
table.sheet.grid td { outline: 0.5px solid #c0c0c0; outline-offset: -0.5px; }
table.sheet th { background: #f2f2f2; border: 0.5px solid #bfbfbf; font-weight: normal; font-size: 9pt; }
.bw td, .bw th { color: #000 !important; background: transparent !important; }
</style></head><body>${body.join('')}</body></html>`
  return { html, widthIn: first?.w ?? 8.27, heightIn: first?.h ?? 11.69, pageCount: total }
}

/** Kept for callers that only need the document HTML. */
export function buildPrintHtml(o: PrintOptions, preview = false): string | null {
  return buildPrint(o, preview)?.html ?? null
}

/** Prints or exports; returns an error message to show, or null. */
export async function runPrint(o: PrintOptions, mode: 'print' | 'pdf'): Promise<string | null> {
  const doc = buildPrint(o)
  if (!doc) return 'There is nothing to print.'
  if (!window.api?.printDocument) {
    // Browser build: print through a hidden frame.
    const f = document.createElement('iframe')
    f.style.cssText = 'position:fixed;width:0;height:0;border:0'
    document.body.appendChild(f)
    f.srcdoc = doc.html
    await new Promise((r) => (f.onload = r))
    f.contentWindow?.print()
    setTimeout(() => f.remove(), 1000)
    return null
  }
  const res = await window.api.printDocument({ html: doc.html, mode, widthIn: doc.widthIn, heightIn: doc.heightIn, defaultName: store.fileName || 'Book1' })
  return res === true || res === null ? null : res
}

export { DEFAULT_PAPER, paperOf }

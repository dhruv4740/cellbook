// Builds the document File > Print and Export > PDF send to the printer: the sheet's cells with
// their formatting (fonts, fills, borders, alignment, merges, conditional formats), skipping
// hidden rows/columns, laid out for the chosen paper, orientation, margins and scaling.
import { store, SheetState } from '../model/store'
import { Range, normRange, CellStyle, BorderSide, colName } from '../model/types'
import { evaluateCf } from '../format/conditional'

export interface PrintOptions {
  what: 'active' | 'workbook' | 'selection'
  landscape: boolean
  pageSize: 'A4' | 'Letter' | 'Legal'
  margins: 'normal' | 'wide' | 'narrow'
  scaling: 'none' | 'fitSheet' | 'fitColumns'
  gridlines: boolean
  headings: boolean
}
/** The print options a sheet's saved Page Layout settings imply. */
export function optionsFor(sheet: SheetState): PrintOptions {
  const p = sheet.layout.pageSetup ?? {}
  return { ...DEFAULT_PRINT, landscape: !!p.landscape, pageSize: p.pageSize ?? 'A4', margins: p.margins ?? 'normal', scaling: p.fitColumns ? 'fitColumns' : 'none', gridlines: !!p.gridlines, headings: !!p.headings }
}
export const DEFAULT_PRINT: PrintOptions = { what: 'active', landscape: false, pageSize: 'A4', margins: 'normal', scaling: 'none', gridlines: false, headings: false }

const MARGINS = { normal: '0.75in 0.7in', wide: '1in 1in', narrow: '0.75in 0.25in' }
const PAGE_W_IN = { A4: 8.27, Letter: 8.5, Legal: 8.5 }
const PAGE_H_IN = { A4: 11.69, Letter: 11, Legal: 14 }
const SIDE_IN = { normal: 0.7, wide: 1, narrow: 0.25 }

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
function printRange(sheet: SheetState, what: PrintOptions['what']): Range | null {
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

function sheetTable(sheet: SheetState, rg: Range, o: PrintOptions): { html: string; widthPx: number } {
  const l = sheet.layout
  const rows: number[] = []
  const cols: number[] = []
  for (let r = rg.r1; r <= rg.r2; r++) if (!l.hiddenRows[r]) rows.push(r)
  for (let c = rg.c1; c <= rg.c2; c++) if (!l.hiddenCols[c]) cols.push(c)
  const merges = l.merges.map(normRange)
  const covered = new Set<string>()
  const span = new Map<string, { rs: number; cs: number }>()
  for (const m of merges) {
    if (m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2) continue
    const rs = rows.filter((r) => r >= m.r1 && r <= m.r2).length
    const cs = cols.filter((c) => c >= m.c1 && c <= m.c2).length
    span.set(`${m.r1},${m.c1}`, { rs, cs })
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`)
  }
  const widthPx = cols.reduce((a, c) => a + store.colWidth(sheet, c), 0) + (o.headings ? 40 : 0)
  const out: string[] = []
  out.push(`<table class="sheet${o.gridlines ? ' grid' : ''}"><colgroup>${o.headings ? '<col style="width:40px">' : ''}${cols.map((c) => `<col style="width:${store.colWidth(sheet, c)}px">`).join('')}</colgroup>`)
  if (o.headings) out.push(`<thead><tr><th></th>${cols.map((c) => `<th>${colName(c)}</th>`).join('')}</tr></thead>`)
  out.push('<tbody>')
  for (const r of rows) {
    out.push(`<tr style="height:${store.rowHeight(sheet, r)}px">${o.headings ? `<th>${r + 1}</th>` : ''}`)
    for (const c of cols) {
      const k = `${r},${c}`
      if (covered.has(k)) continue
      const sp = span.get(k)
      const base: CellStyle | undefined = store.getStyle(sheet, r, c)
      const cf = evaluateCf(store, sheet, r, c)
      const st: CellStyle = { ...base, ...(cf.style ?? {}) }
      const d = store.getDisplay(sheet, r, c)
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
      const h = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'centerContinuous' ? 'center' : st.hAlign === 'distributed' ? 'justify' : st.hAlign) : d.align ?? 'left'
      css.push(`text-align:${h}`)
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
      const attrs = (sp && sp.rs > 1 ? ` rowspan="${sp.rs}"` : '') + (sp && sp.cs > 1 ? ` colspan="${sp.cs}"` : '')
      out.push(`<td${attrs}${css.length ? ` style="${css.join(';')}"` : ''}>${esc(d.text)}</td>`)
    }
    out.push('</tr>')
  }
  out.push('</tbody></table>')
  return { html: out.join(''), widthPx }
}

/** A complete HTML document for printing / PDF, or null when there is nothing to print. */
export function buildPrintHtml(o: PrintOptions): string | null {
  const sheets = o.what === 'workbook' ? store.sheets.filter((s) => !s.layout.sheetHidden) : [store.active]
  const parts: string[] = []
  let maxW = 0
  for (const sh of sheets) {
    const rg = printRange(sh, o.what)
    if (!rg) continue
    const t = sheetTable(sh, rg, o)
    maxW = Math.max(maxW, t.widthPx)
    parts.push(`<section>${t.html}</section>`)
  }
  if (!parts.length) return null
  const pageWIn = (o.landscape ? PAGE_H_IN : PAGE_W_IN)[o.pageSize] - 2 * SIDE_IN[o.margins]
  const pageWPx = pageWIn * 96
  const zoom = o.scaling !== 'none' && maxW > pageWPx ? pageWPx / maxW : 1
  return `<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: ${o.pageSize} ${o.landscape ? 'landscape' : 'portrait'}; margin: ${MARGINS[o.margins]}; }
html, body { margin: 0; }
body { font: 11pt Calibri, 'Segoe UI', Arial, sans-serif; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
section { zoom: ${zoom.toFixed(4)}; }
section + section { break-before: page; }
table.sheet { border-collapse: collapse; table-layout: fixed; }
table.sheet td { padding: 0 3px; overflow: hidden; white-space: pre; line-height: 1.15; }
table.sheet.grid td { outline: 0.5px solid #c0c0c0; outline-offset: -0.5px; }
table.sheet th { background: #f2f2f2; border: 0.5px solid #bfbfbf; font-weight: normal; font-size: 9pt; }
thead { display: table-header-group; }
tr { break-inside: avoid; }
</style></head><body>${parts.join('')}</body></html>`
}

/** Prints or exports; returns an error message to show, or null. */
export async function runPrint(o: PrintOptions, mode: 'print' | 'pdf'): Promise<string | null> {
  const html = buildPrintHtml(o)
  if (!html) return 'There is nothing to print.'
  if (!window.api?.printDocument) {
    // Browser build: print through a hidden frame.
    const f = document.createElement('iframe')
    f.style.cssText = 'position:fixed;width:0;height:0;border:0'
    document.body.appendChild(f)
    f.srcdoc = html
    await new Promise((r) => (f.onload = r))
    f.contentWindow?.print()
    setTimeout(() => f.remove(), 1000)
    return null
  }
  const res = await window.api.printDocument({ html, mode, landscape: o.landscape, pageSize: o.pageSize, defaultName: store.fileName || 'Book1' })
  return res === true || res === null ? null : res
}

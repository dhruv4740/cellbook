// Reads and writes the worksheet elements behind Page Layout: pageSetup, pageMargins, printOptions,
// headerFooter, rowBreaks/colBreaks (and sheetPr/pageSetUpPr for fit-to-page). Element order follows the schema.
import { PageSetup } from '../model/types'
import { DEFAULT_PAPER, marginPresetOf } from '../model/pageSetup'
import { parseAttrs, decodeXmlEntities } from './fastXlsxXml'

const yes = (v: string | undefined): boolean => v === '1' || v === 'true'

export function parsePageSetupXml(xml: string): PageSetup | undefined {
  const tail = xml.slice(Math.max(0, xml.lastIndexOf('</sheetData>')))
  const out: PageSetup = {}
  const attrsOf = (re: RegExp): Record<string, string> | undefined => {
    const m = re.exec(tail)
    return m ? parseAttrs(m[1], 0, m[1].length) : undefined
  }
  const ps = attrsOf(/<pageSetup\b([^>]*?)\/?>/)
  if (ps) {
    if (ps.orientation === 'landscape') out.landscape = true
    if (ps.paperSize) out.paperSize = parseInt(ps.paperSize, 10)
    if (ps.scale && parseInt(ps.scale, 10) !== 100) out.scale = parseInt(ps.scale, 10)
    if (ps.fitToWidth !== undefined) out.fitToWidth = parseInt(ps.fitToWidth, 10)
    if (ps.fitToHeight !== undefined) out.fitToHeight = parseInt(ps.fitToHeight, 10)
    if (ps.firstPageNumber && yes(ps.useFirstPageNumber)) out.firstPageNumber = parseInt(ps.firstPageNumber, 10)
    if (ps.pageOrder === 'overThenDown') out.pageOrder = 'overThenDown'
    if (yes(ps.blackAndWhite)) out.blackAndWhite = true
    if (yes(ps.draft)) out.draft = true
    if (ps.cellComments === 'atEnd' || ps.cellComments === 'asDisplayed') out.comments = ps.cellComments
    if (ps.errors === 'blank' || ps.errors === 'dash' || ps.errors === 'NA') out.errors = ps.errors
  }
  const head = xml.slice(0, Math.max(0, xml.indexOf('<sheetData')))
  if (/<pageSetUpPr\b[^>]*fitToPage="(1|true)"/.test(head)) {
    out.fitToPage = true
    out.fitToWidth ??= 1
    out.fitToHeight ??= 1
  }
  const po = attrsOf(/<printOptions\b([^>]*?)\/?>/)
  if (po) {
    if (yes(po.gridLines)) out.gridlines = true
    if (yes(po.headings)) out.headings = true
    if (yes(po.horizontalCentered)) out.centerH = true
    if (yes(po.verticalCentered)) out.centerV = true
  }
  const pm = attrsOf(/<pageMargins\b([^>]*?)\/?>/)
  if (pm) {
    const m: NonNullable<PageSetup['margins']> = {}
    for (const k of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) if (pm[k] !== undefined) m[k] = parseFloat(pm[k])
    out.margins = m
  }
  const hfM = /<headerFooter\b([^>]*?)(?:\/>|>([\s\S]*?)<\/headerFooter>)/.exec(tail)
  if (hfM) {
    const a = parseAttrs(hfM[1], 0, hfM[1].length)
    const hf: NonNullable<PageSetup['hf']> = {}
    if (yes(a.differentOddEven)) hf.differentOddEven = true
    if (yes(a.differentFirst)) hf.differentFirst = true
    if (a.scaleWithDoc === '0' || a.scaleWithDoc === 'false') hf.scaleWithDoc = false
    if (a.alignWithMargins === '0' || a.alignWithMargins === 'false') hf.alignWithMargins = false
    for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const) {
      const t = new RegExp(`<${k}>([\\s\\S]*?)</${k}>`).exec(hfM[2] ?? '')
      if (t) hf[k] = decodeXmlEntities(t[1])
    }
    if (Object.keys(hf).length) out.hf = hf
  }
  const breaks = (tag: 'rowBreaks' | 'colBreaks'): number[] | undefined => {
    const m = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(tail)
    if (!m) return undefined
    const ids = [...m[1].matchAll(/<brk\b([^>]*?)\/?>/g)].map((b) => parseAttrs(b[1], 0, b[1].length)).filter((a) => a.man !== '0').map((a) => parseInt(a.id, 10)).filter((n) => n > 0)
    return ids.length ? ids : undefined
  }
  const rb = breaks('rowBreaks'), cb = breaks('colBreaks')
  if (rb) out.rowBreaks = rb
  if (cb) out.colBreaks = cb
  // Keep the model minimal: settings equal to the defaults are not stored.
  if (out.margins && marginPresetOf(out) === 'normal' && Object.keys(out.margins).length === 6) delete out.margins
  if (!out.fitToPage) { delete out.fitToWidth; delete out.fitToHeight }
  if (out.paperSize === DEFAULT_PAPER) delete out.paperSize
  return Object.keys(out).length ? out : undefined
}

const esc = (t: string): string => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const num = (n: number): string => String(+n.toFixed(6))

/** The printOptions .. colBreaks run of elements, in schema order. */
export function buildPageSetupXml(p: PageSetup): string {
  const m = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3, ...(p.margins ?? {}) }
  let x = ''
  const po = (p.gridlines ? ' gridLines="1"' : '') + (p.headings ? ' headings="1"' : '') + (p.centerH ? ' horizontalCentered="1"' : '') + (p.centerV ? ' verticalCentered="1"' : '')
  if (po) x += `<printOptions${po}/>`
  x += `<pageMargins left="${num(m.left)}" right="${num(m.right)}" top="${num(m.top)}" bottom="${num(m.bottom)}" header="${num(m.header)}" footer="${num(m.footer)}"/>`
  let a = ''
  a += ` paperSize="${p.paperSize ?? 9}"`
  if (p.scale && p.scale !== 100) a += ` scale="${p.scale}"`
  if (p.fitToPage) { if ((p.fitToWidth ?? 1) !== 1) a += ` fitToWidth="${p.fitToWidth}"`; if ((p.fitToHeight ?? 1) !== 1) a += ` fitToHeight="${p.fitToHeight}"` }
  if (p.firstPageNumber !== undefined) a += ` firstPageNumber="${p.firstPageNumber}"`
  if (p.pageOrder === 'overThenDown') a += ' pageOrder="overThenDown"'
  a += ` orientation="${p.landscape ? 'landscape' : 'portrait'}"`
  if (p.blackAndWhite) a += ' blackAndWhite="1"'
  if (p.draft) a += ' draft="1"'
  if (p.comments) a += ` cellComments="${p.comments}"`
  if (p.firstPageNumber !== undefined) a += ' useFirstPageNumber="1"'
  if (p.errors) a += ` errors="${p.errors}"`
  x += `<pageSetup${a}/>`
  const hf = p.hf
  if (hf) {
    let ha = ''
    if (hf.differentOddEven) ha += ' differentOddEven="1"'
    if (hf.differentFirst) ha += ' differentFirst="1"'
    if (hf.scaleWithDoc === false) ha += ' scaleWithDoc="0"'
    if (hf.alignWithMargins === false) ha += ' alignWithMargins="0"'
    let kids = ''
    for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const) if (hf[k]) kids += `<${k}>${esc(hf[k]!)}</${k}>`
    if (ha || kids) x += `<headerFooter${ha}>${kids}</headerFooter>`
  }
  const brk = (tag: string, ids: number[] | undefined, max: number): string =>
    ids && ids.length ? `<${tag} count="${ids.length}" manualBreakCount="${ids.length}">${[...ids].sort((u, v) => u - v).map((id) => `<brk id="${id}" max="${max}" man="1"/>`).join('')}</${tag}>` : ''
  x += brk('rowBreaks', p.rowBreaks, 16383)
  x += brk('colBreaks', p.colBreaks, 1048575)
  return x
}

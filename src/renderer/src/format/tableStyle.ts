// Excel tables: the look of a table style (header row, banded rows, first / last column) painted under the cells.
// Approximates the built-in TableStyleLight / Medium / Dark families with the Office accent colours.
import type { CellStyle } from '../model/types'
import type { SheetState } from '../model/store'

const ACCENTS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47']
const mix = (hex: string, t: number, to = 255): string => {
  const n = parseInt(hex.slice(1), 16)
  const ch = (v: number): number => Math.round(v + (to - v) * t)
  return '#' + [ch((n >> 16) & 255), ch((n >> 8) & 255), ch(n & 255)].map((x) => x.toString(16).padStart(2, '0')).join('')
}

interface Look { header?: Partial<CellStyle>; band?: string }

function lookOf(style: string): Look {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(style)
  if (!m) return { band: '#f2f2f2' }
  const n = +m[2]
  const kind = m[1]
  const accent = (k: number): string => ACCENTS[((k % 6) + 6) % 6]
  if (kind === 'Medium') {
    const coloured = (n >= 2 && n <= 7) || (n >= 9 && n <= 14)
    if (!coloured) return { header: { fill: '#000000', color: '#ffffff', bold: true }, band: '#d9d9d9' }
    const a = accent(n <= 7 ? n - 2 : n - 9)
    return { header: { fill: a, color: '#ffffff', bold: true }, band: mix(a, 0.8) }
  }
  if (kind === 'Light') {
    if (n === 1 || n === 15) return { header: { bold: true }, band: '#d9d9d9' }
    if (n >= 2 && n <= 7) return { header: { bold: true }, band: mix(accent(n - 2), 0.8) }
    if (n >= 8 && n <= 14) return { header: { fill: accent(n - 8), color: '#ffffff', bold: true }, band: mix(accent(n - 8), 0.8) }
    return { header: { bold: true }, band: '#f2f2f2' }
  }
  if (n === 1) return { header: { fill: '#404040', color: '#ffffff', bold: true }, band: '#a6a6a6' }
  const a = accent(n - 2)
  return { header: { fill: mix(a, 0.5, 0), color: '#ffffff', bold: true }, band: mix(a, 0.6) }
}

const cache = new Map<string, Look>()

/** The table-style formatting for a cell inside one of the sheet's tables, or null. */
export function tableStyleFor(sheet: SheetState, r: number, c: number): Partial<CellStyle> | null {
  const tables = sheet.layout.tables
  if (!tables?.length) return null
  for (const t of tables) {
    if (!t.style || r < t.r1 || r > t.r2 || c < t.c1 || c > t.c2) continue
    let look = cache.get(t.style)
    if (!look) { look = lookOf(t.style); cache.set(t.style, look) }
    const out: Partial<CellStyle> = {}
    if (t.headerRows && r < t.r1 + t.headerRows) Object.assign(out, look.header)
    else if (t.rowStripes && look.band && r - t.r1 - t.headerRows >= 0 && (r - t.r1 - t.headerRows) % 2 === 0 && r <= t.r2 - t.totalsRows) out.fill = look.band
    if ((t.firstCol && c === t.c1) || (t.lastCol && c === t.c2)) out.bold = true
    return Object.keys(out).length ? out : null
  }
  return null
}

// Excel tables (ListObjects) are not modelled in Cellbook, so formulas that use structured references
// (Orders[Qty], Orders[[#This Row],[Qty]], Orders[#All] ...) are rewritten to plain ranges on load.
// Without this they evaluate to #ERROR! and, once saved, Excel refuses to open the file.
import { colName } from '../model/types'

export interface TableDef {
  name: string
  sheet: string
  r1: number
  c1: number
  r2: number
  c2: number
  headerRows: number
  totalsRows: number
  columns: string[]
}

const quoteSheet = (s: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`)
const unescape = (s: string): string => s.replace(/'(.)/g, '$1')

/** Splits "[#All],[Qty]" / "[Qty]:[Rate]" into bracketed items, honouring Excel's ' escapes. */
function items(spec: string): string[] {
  const out: string[] = []
  for (let i = 0; i < spec.length; i++) {
    if (spec[i] !== '[') continue
    let j = i + 1
    while (j < spec.length && spec[j] !== ']') j += spec[j] === "'" ? 2 : 1
    out.push(spec.slice(i + 1, j))
    if (spec[j + 1] === ':') out.push(':')
    i = j
  }
  return out
}

function toRange(t: TableDef, spec: string, row: number, formulaSheet: string): string | null {
  const s = spec.trim()
  const parts = s === '' ? [] : s.startsWith('[') ? items(s) : [s]
  const areas = parts.filter((p) => p.startsWith('#')).map((p) => p.toLowerCase())
  const cols = parts.filter((p) => p !== ':' && !p.startsWith('#')).map((p) => unescape(p.startsWith('@') ? p.slice(1) : p).trim())
  if (s.startsWith('@')) areas.push('#this row')
  const dataTop = t.r1 + t.headerRows
  const dataBottom = t.r2 - t.totalsRows
  let top = dataTop
  let bottom = dataBottom
  if (areas.includes('#this row')) top = bottom = row
  else if (areas.includes('#all')) { top = t.r1; bottom = t.r2 }
  else {
    const hdr = areas.includes('#headers'), tot = areas.includes('#totals'), data = areas.includes('#data') || (!hdr && !tot)
    top = hdr ? t.r1 : data ? dataTop : dataBottom + 1
    bottom = tot ? t.r2 : data ? dataBottom : t.r1
  }
  let c1 = t.c1
  let c2 = t.c2
  if (cols.length) {
    const idx = cols.map((c) => t.columns.findIndex((n) => n.toLowerCase() === c.toLowerCase()))
    if (idx.some((i) => i < 0)) return null
    c1 = t.c1 + Math.min(...idx)
    c2 = t.c1 + Math.max(...idx)
  }
  const sheet = t.sheet === formulaSheet ? '' : quoteSheet(t.sheet) + '!'
  const a = `$${colName(c1)}$${top + 1}`
  const b = `$${colName(c2)}$${bottom + 1}`
  return sheet + (a === b ? a : `${a}:${b}`)
}

/** Rewrites structured references in a formula (without its leading "=") for a cell at `row` on `sheet`. */
export function resolveStructuredRefs(formula: string, tables: TableDef[], sheet: string, row: number): string {
  if (!tables.length || formula.indexOf('[') === -1) return formula
  const byName = new Map(tables.map((t) => [t.name.toLowerCase(), t]))
  let out = ''
  let i = 0
  while (i < formula.length) {
    const ch = formula[i]
    if (ch === '"') {
      const end = formula.indexOf('"', i + 1)
      const stop = end < 0 ? formula.length : end + 1
      out += formula.slice(i, stop)
      i = stop
      continue
    }
    const m = /^[A-Za-z_\\][A-Za-z0-9_.\\]*\[/.exec(formula.slice(i))
    const prev = formula[i - 1]
    if (m && !(prev && /[A-Za-z0-9_.!']/.test(prev))) {
      const t = byName.get(m[0].slice(0, -1).toLowerCase())
      if (t) {
        let depth = 0
        let j = i + m[0].length - 1
        for (; j < formula.length; j++) {
          if (formula[j] === "'") { j++; continue }
          if (formula[j] === '[') depth++
          else if (formula[j] === ']' && --depth === 0) break
        }
        const ref = toRange(t, formula.slice(i + m[0].length, j), row, sheet)
        if (ref) {
          out += ref
          i = j + 1
          continue
        }
      }
    }
    out += ch
    i++
  }
  return out
}

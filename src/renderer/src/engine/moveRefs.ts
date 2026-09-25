// Cut + Paste in Excel moves cells, and every formula that pointed at a moved cell follows it
// (absolute or relative alike). A range ref follows only when the whole range was moved.
import { colName, colIndex } from '../model/types'
import type { Range } from '../model/types'

const RANGE_RE =
  /((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/g

function sheetOf(prefix: string): string | null {
  if (!prefix) return null
  const p = prefix.slice(0, -1)
  return p.startsWith("'") ? p.slice(1, -1).replace(/''/g, "'") : p
}

/** `formulaSheet`: sheet holding the formula; `srcSheet`: sheet the cut cells came from. */
export function moveRefs(formula: string, formulaSheet: string, srcSheet: string, src: Range, dr: number, dc: number, destSheet: string): string {
  if (!formula.startsWith('=')) return formula
  const inside = (r: number, c: number): boolean => r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2
  const parts = formula.split(/("(?:[^"]|"")*")/g)
  let changed = false
  for (let p = 0; p < parts.length; p += 2) {
    parts[p] = parts[p].replace(RANGE_RE, (full, prefix = '', ca1, cs1, ra1, rs1, ca2, cs2, ra2, rs2, off: number, str: string) => {
      const before = str[off - 1]
      const after = str[off + full.length]
      if ((before && /[A-Za-z0-9_$.]/.test(before)) || after === '(' || (after && /[A-Za-z0-9_]/.test(after))) return full
      const refSheet = sheetOf(prefix) ?? formulaSheet
      if (refSheet !== srcSheet) return full
      const r1 = +rs1 - 1, c1 = colIndex(cs1)
      const r2 = rs2 ? +rs2 - 1 : r1, c2 = cs2 ? colIndex(cs2) : c1
      if (!inside(r1, c1) || !inside(r2, c2)) return full
      changed = true
      const needPrefix = prefix || destSheet !== formulaSheet
      const pre = destSheet === srcSheet ? prefix : needPrefix ? (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(destSheet) ? destSheet : `'${destSheet.replace(/'/g, "''")}'`) + '!' : ''
      const a = `${ca1}${colName(c1 + dc)}${ra1}${r1 + dr + 1}`
      return pre + a + (rs2 ? `:${ca2}${colName(c2 + dc)}${ra2}${r2 + dr + 1}` : '')
    })
  }
  return changed ? parts.join('') : formula
}

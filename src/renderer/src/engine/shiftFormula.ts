// Shifts relative cell references inside a formula by (dr, dc) rows/columns, as Excel does when a
// formula is copied, pasted, filled, or when rows/columns it points at move because of an edit.
// Absolute references ($A$1) and absolute parts (A$1, $A1) are preserved. Sheet-qualified refs
// ('Sheet2'!A1, Sheet2!A1) are shifted the same way. References inside string literals are left
// untouched. A reference that would land outside the sheet becomes #REF!.
import { MAX_ROWS, MAX_COLS, colName, colIndex } from '../model/types'

// sheet-prefix (quoted or bare) + optional $ + 1-3 letters + optional $ + 1-7 digits
const REF_RE = /((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g

function isWordChar(ch: string | undefined): boolean {
  return !!ch && /[A-Za-z0-9_$]/.test(ch)
}

/** Splits a formula into alternating [outside-string, inside-string, outside-string, ...] parts. */
function splitStrings(formula: string): string[] {
  return formula.split(/("(?:[^"]|"")*")/g)
}

export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!formula.startsWith('=')) return formula
  if (dr === 0 && dc === 0) return formula
  const parts = splitStrings(formula)
  for (let p = 0; p < parts.length; p += 2) {
    const src = parts[p]
    let out = ''
    let last = 0
    REF_RE.lastIndex = 0
    let m: RegExpExecArray | null
    while ((m = REF_RE.exec(src))) {
      const [full, sheetPrefix = '', colAbs, colStr, rowAbs, rowStr] = m
      const before = src[m.index - 1]
      const after = src[m.index + full.length]
      // Skip function-call-like names (FOO() ) and identifiers that merely look like a ref.
      if (isWordChar(before) || after === '(') continue
      let col = colIndex(colStr)
      let row = +rowStr - 1
      if (colAbs !== '$') col += dc
      if (rowAbs !== '$') row += dr
      out += src.slice(last, m.index)
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        out += '#REF!'
      } else {
        out += sheetPrefix + colAbs + colName(col) + rowAbs + (row + 1)
      }
      last = m.index + full.length
    }
    out += src.slice(last)
    parts[p] = out
  }
  return parts.join('')
}

// Data > Advanced Filter: criteria-range filtering. Criteria rows are OR'd, cells within a row AND'd; text criteria
// match cells that BEGIN with the text (Excel's rule), "=x" matches exactly, and * ? wildcards work.
import { CellValue } from '../model/types'

export interface CellRead { text: string; value: CellValue }

function wildcardRegex(pat: string, anchored: boolean): RegExp {
  const src = pat.replace(/~([*?])|([.+^${}()|[\]\\])|(\*)|(\?)/g, (_m, esc, sp, star, q) => (esc ? '\\' + esc : sp ? '\\' + sp : star ? '.*' : q ? '.' : ''))
  return new RegExp((anchored ? '^' : '^') + src + (anchored ? '$' : ''), 'i')
}

/** Does a cell satisfy one criteria cell (e.g. ">5", "=abc", "Da*", "<>")? */
export function matchesCriterion(criterion: string, cell: CellRead): boolean {
  const c = criterion.trim()
  if (c === '') return true
  const m = /^(>=|<=|<>|>|<|=)?(.*)$/s.exec(c)!
  const op = m[1] ?? ''
  const operand = m[2]
  const blank = cell.value === null || cell.value === undefined || cell.value === ''
  const cellNum = typeof cell.value === 'number' ? cell.value : null
  const opNum = operand.trim() !== '' && !isNaN(Number(operand)) ? Number(operand) : null
  if (operand === '' ) {
    if (op === '=') return blank
    if (op === '<>') return !blank
    return true
  }
  if (opNum !== null && (op !== '' || true)) {
    if (cellNum === null) return op === '<>' // text never equals a number
    switch (op) {
      case '>': return cellNum > opNum
      case '<': return cellNum < opNum
      case '>=': return cellNum >= opNum
      case '<=': return cellNum <= opNum
      case '<>': return cellNum !== opNum
      default: return cellNum === opNum
    }
  }
  const t = cell.text
  if (op === '=') return wildcardRegex(operand, true).test(t)
  if (op === '<>') return !wildcardRegex(operand, true).test(t)
  if (op === '') return !blank && wildcardRegex(operand, false).test(t)
  // ordered text comparison
  const cmp = t.toLowerCase().localeCompare(operand.toLowerCase())
  return op === '>' ? cmp > 0 : op === '<' ? cmp < 0 : op === '>=' ? cmp >= 0 : cmp <= 0
}

/**
 * Rows (r1+1..r2 of `list`) that pass the criteria block. `criteria` is a grid whose first row holds column titles
 * matched to the list's titles; the rows below are alternatives.
 */
export function advancedFilterRows(
  listHeaders: string[],
  rowCount: number,
  read: (row: number, col: number) => CellRead,
  criteria: string[][]
): number[] {
  if (criteria.length < 1) return Array.from({ length: rowCount }, (_, i) => i)
  const titles = criteria[0].map((t) => t.trim().toLowerCase())
  const colOf = titles.map((t) => (t === '' ? -1 : listHeaders.findIndex((h) => h.trim().toLowerCase() === t)))
  const rules = criteria.slice(1).map((row) => row.map((cr, i) => ({ col: colOf[i], cr })).filter((x) => x.col >= 0 && x.cr.trim() !== ''))
  const out: number[] = []
  for (let i = 0; i < rowCount; i++) {
    // No criteria rows (or an all-blank one) means everything passes.
    const ok = rules.length === 0 || rules.some((rule) => rule.every((x) => matchesCriterion(x.cr, read(i, x.col))))
    if (ok) out.push(i)
  }
  return out
}

/** Keeps the first of each set of rows with identical (case-insensitive) text. */
export function uniqueRows(rows: number[], key: (row: number) => string): number[] {
  const seen = new Set<string>()
  return rows.filter((r) => { const k = key(r).toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true })
}

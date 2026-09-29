// Formulas > Evaluate Formula: steps through a formula the way Excel's dialog does. Each step replaces the underlined
// part (the next cell reference, or else the innermost function call / parenthesised group) with its value.
import { DetailedCellError } from 'hyperformula'
import { store, SheetState } from '../model/store'
import { parseA1 } from '../model/types'

const REF = /(?<![\w.$!'])((?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(!])/g
const CALL = /(?:[A-Za-z_][\w.]*)?\(([^()]*)\)/g

/** The expression with string literals blanked out, so scans never look inside them. */
function masked(expr: string): string {
  let out = ''
  let inStr = false
  for (let i = 0; i < expr.length; i++) {
    const ch = expr[i]
    if (ch === '"') { if (inStr && expr[i + 1] === '"') { out += '  '; i++; continue } inStr = !inStr; out += ' '; continue }
    out += inStr ? ' ' : ch
  }
  return out
}

const literal = (v: unknown): string => {
  if (v === null || v === undefined || v === '') return '""'
  if (typeof v === 'number') return String(v)
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (v instanceof DetailedCellError) return String(v.value)
  if (typeof v === 'object' && 'error' in (v as object)) return String((v as { error: unknown }).error)
  return '"' + String(v).replace(/"/g, '""') + '"'
}

function refValue(sheet: SheetState, ref: string): string {
  let target = sheet
  let a1 = ref
  const bang = ref.lastIndexOf('!')
  if (bang >= 0) {
    const name = ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
    const s = store.sheets.find((x) => x.name.toLowerCase() === name.toLowerCase())
    if (!s) return '#REF!'
    target = s
    a1 = ref.slice(bang + 1)
  }
  const rg = parseA1(a1.replace(/\$/g, ''))
  if (!rg) return '#REF!'
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return literal(store.getValue(target, rg.r1, rg.c1))
  const rows: string[] = []
  for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 199); r++) {
    const cols: string[] = []
    for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 49); c++) cols.push(literal(store.getValue(target, r, c)))
    rows.push(cols.join(','))
  }
  return '{' + rows.join(';') + '}'
}

export interface EvalStep { start: number; end: number; replacement: string }

/** The next step for `expr` (no leading "="), or null when it is already a single value. */
export function nextStep(expr: string, sheet: SheetState): EvalStep | null {
  const m = masked(expr)
  REF.lastIndex = 0
  const ref = REF.exec(m)
  if (ref) return { start: ref.index, end: ref.index + ref[0].length, replacement: refValue(sheet, expr.slice(ref.index, ref.index + ref[0].length)) }
  CALL.lastIndex = 0
  let last: RegExpExecArray | null = null
  for (let x = CALL.exec(m); x; x = CALL.exec(m)) { last = x; break }
  const target = last ? { start: last.index, end: last.index + last[0].length } : { start: 0, end: expr.length }
  const piece = expr.slice(target.start, target.end)
  if (!last && /^\s*(?:-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?|"(?:[^"]|"")*"|TRUE|FALSE|#[A-Z/0!?]+)\s*$/i.test(piece)) return null
  let value: unknown
  try { value = store.hf.calculateFormula('=' + piece, store.hfId(sheet)) } catch { return null }
  const rep = Array.isArray(value) ? '{' + (value as unknown[][]).map((row) => row.map(literal).join(',')).join(';') + '}' : literal(value)
  if (rep === piece.trim()) return null
  return { start: target.start, end: target.end, replacement: rep }
}

export function applyStep(expr: string, s: EvalStep): string {
  return expr.slice(0, s.start) + s.replacement + expr.slice(s.end)
}

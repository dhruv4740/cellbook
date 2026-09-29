// Excel's green "error checking" triangles: numbers stored as text, and formulas that evaluate to an error.
import { store, SheetState } from '../model/store'
import { MAX_COLS } from '../model/types'

const NUMBER_TEXT = /^\s*[-+]?(?:\d+(?:,\d{3})*|\d*)(?:\.\d+)?(?:[eE][-+]?\d+)?%?\s*$/

export function cellErrorKind(sheet: SheetState, r: number, c: number): string | null {
  const v = store.getValue(sheet, r, c)
  if (typeof v === 'string') {
    if (v.length === 0 || v.length > 40) return null
    const ch = v.charCodeAt(0)
    // digits, space, sign, '.' only: skip the regex for ordinary words
    if (!((ch >= 48 && ch <= 57) || ch === 32 || ch === 43 || ch === 45 || ch === 46)) return null
    return /\d/.test(v) && NUMBER_TEXT.test(v) ? 'Number Stored as Text' : null
  }
  if (v && typeof v === 'object' && 'error' in v) return store.getRaw(sheet, r, c).startsWith('=') ? `Formula results in an error (${String((v as { error: unknown }).error)})` : null
  return null
}

/** Formulas > Error Checking: the next flagged cell after the active one (wrapping), or null. */
export function nextErrorCell(sheet: SheetState): { r: number; c: number; kind: string } | null {
  const used = store.usedRange(sheet)
  if (!used) return null
  const a = sheet.selection.active
  const total = (used.r2 + 1) * (used.c2 + 1)
  const start = a.r * (used.c2 + 1) + a.c + 1
  for (let i = 0; i < Math.min(total, MAX_COLS * 100000); i++) {
    const idx = (start + i) % total
    const r = Math.floor(idx / (used.c2 + 1)), c = idx % (used.c2 + 1)
    const kind = cellErrorKind(sheet, r, c)
    if (kind) return { r, c, kind }
  }
  return null
}

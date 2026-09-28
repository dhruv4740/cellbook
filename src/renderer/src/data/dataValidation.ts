import { WorkbookStore, SheetState } from '../model/store'
import {
  DataValidationRule,
  DvErrorStyle,
  DvOperator,
  Range,
  inRange,
  normRange,
  parseA1
} from '../model/types'

export interface ValidationResult {
  ok: boolean
  style?: DvErrorStyle
  title?: string
  message?: string
}

/** Finds the first data validation rule that covers cell (r, c). */
export function getValidationForCell(
  sheet: SheetState,
  r: number,
  c: number
): DataValidationRule | undefined {
  const rules = sheet.layout.dataValidations
  if (!rules || !rules.length) return undefined

  for (const rule of rules) {
    for (const rg of rule.ranges) {
      if (inRange(rg, r, c)) return rule
    }
  }
  return undefined
}

/** Extracts allowed list items for a list validation rule. */
export function getListItems(
  store: WorkbookStore,
  sheet: SheetState,
  rule: DataValidationRule
): string[] {
  if (rule.allow !== 'list' || !rule.formula1) return []
  const src = rule.formula1.trim()

  if (src.startsWith('=')) {
    // Range reference like =A1:A5 or =Sheet1!A1:A5
    const refStr = src.slice(1).replace(/\$/g, '')
    const parsed = parseA1(refStr)
    if (parsed) {
      const items: string[] = []
      const n = normRange(parsed)
      for (let r = n.r1; r <= n.r2; r++) {
        for (let c = n.c1; c <= n.c2; c++) {
          const val = store.getRaw(sheet, r, c).trim()
          if (val) items.push(val)
        }
      }
      return items
    }
  }

  // Comma-separated list like "Apple, Banana, Orange" or 'Apple, Banana, Orange'
  const clean = src.replace(/^"|"$/g, '')
  return clean.split(',').map((s) => s.trim()).filter(Boolean)
}

function checkOperator(val: number, op: DvOperator | undefined, f1: number, f2?: number): boolean {
  if (op === undefined || op === 'between') {
    return val >= f1 && val <= (f2 ?? Infinity)
  }
  if (op === 'notBetween') {
    return val < f1 || val > (f2 ?? -Infinity)
  }
  if (op === 'equal') return val === f1
  if (op === 'notEqual') return val !== f1
  if (op === 'greaterThan') return val > f1
  if (op === 'lessThan') return val < f1
  if (op === 'greaterThanOrEqual') return val >= f1
  if (op === 'lessThanOrEqual') return val <= f1
  return true
}

/** Validates input entered into a cell against data validation rules. */
export function validateInput(
  store: WorkbookStore,
  sheet: SheetState,
  r: number,
  c: number,
  raw: string
): ValidationResult {
  const rule = getValidationForCell(sheet, r, c)
  if (!rule || rule.allow === 'any') return { ok: true }

  const trimmed = raw.trim()

  // Blank check
  if (trimmed === '') {
    if (rule.ignoreBlank !== false) return { ok: true }
    return fail(rule)
  }

  const f1Num = Number(rule.formula1 ?? 0)
  const f2Num = rule.formula2 !== undefined ? Number(rule.formula2) : undefined

  if (rule.allow === 'whole') {
    if (!/^-?\d+$/.test(trimmed)) return fail(rule)
    const n = Number(trimmed)
    if (!checkOperator(n, rule.operator, f1Num, f2Num)) return fail(rule)
    return { ok: true }
  }

  if (rule.allow === 'decimal') {
    if (isNaN(Number(trimmed))) return fail(rule)
    const n = Number(trimmed)
    if (!checkOperator(n, rule.operator, f1Num, f2Num)) return fail(rule)
    return { ok: true }
  }

  if (rule.allow === 'textLength') {
    const len = trimmed.length
    if (!checkOperator(len, rule.operator, f1Num, f2Num)) return fail(rule)
    return { ok: true }
  }

  if (rule.allow === 'list') {
    const items = getListItems(store, sheet, rule)
    const match = items.some((item) => item.toLowerCase() === trimmed.toLowerCase())
    if (!match) return fail(rule)
    return { ok: true }
  }

  if (rule.allow === 'date' || rule.allow === 'time') {
    const parsed = Date.parse(trimmed)
    if (isNaN(parsed)) return fail(rule)
    return { ok: true }
  }

  if (rule.allow === 'custom') {
    if (rule.formula1) {
      try {
        const formula = rule.formula1.startsWith('=') ? rule.formula1.slice(1) : rule.formula1
        const res = store.hf.calculateFormula(formula, store.hfId(sheet))
        if (res === false || res === 0) return fail(rule)
      } catch {
        // Fallback
      }
    }
    return { ok: true }
  }

  return { ok: true }
}

function fail(rule: DataValidationRule): ValidationResult {
  return {
    ok: false,
    style: rule.errorStyle || 'stop',
    title: rule.errorTitle || 'Microsoft Excel',
    message:
      rule.errorMessage ||
      "This value doesn't match the data validation restrictions defined for this cell."
  }
}

/** Adds or replaces a data validation rule for ranges. */
export function addDataValidation(
  store: WorkbookStore,
  sheet: SheetState,
  rule: Omit<DataValidationRule, 'id'>
): void {
  if (store.refuseIfProtected(sheet)) return
  store.updateLayout(sheet, 'Data Validation', (l) => {
    const list = [...(l.dataValidations || [])]
    const id = 'dv_' + Math.random().toString(36).slice(2, 9)
    list.push({ ...rule, id })
    l.dataValidations = list
  })
}

/** Clears all data validation rules in ranges. */
export function clearDataValidations(
  store: WorkbookStore,
  sheet: SheetState,
  ranges: Range[]
): void {
  if (store.refuseIfProtected(sheet)) return
  store.updateLayout(sheet, 'Clear Validation', (l) => {
    if (!l.dataValidations) return
    l.dataValidations = l.dataValidations.filter(
      (rule) => !rule.ranges.some((rg) => ranges.some((target) => rangesOverlap(rg, target)))
    )
  })
}

function rangesOverlap(a: Range, b: Range): boolean {
  const na = normRange(a)
  const nb = normRange(b)
  return !(na.r2 < nb.r1 || na.r1 > nb.r2 || na.c2 < nb.c1 || na.c1 > nb.c2)
}

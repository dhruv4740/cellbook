// Excel functions HyperFormula lacks. Registered once, before any HyperFormula instance is built.
import { FunctionPlugin, FunctionArgumentType, HyperFormula, SimpleRangeValue, CellError, ErrorType } from 'hyperformula'
type InterpreterValue = any

function flatten(args: unknown[]): unknown[] {
  const out: unknown[] = []
  for (const a of args) {
    if (a instanceof SimpleRangeValue) for (const row of a.data) for (const v of row) out.push(v)
    else out.push(a)
  }
  return out
}
function toText(v: unknown): string {
  if (v === null || v === undefined || typeof v === 'symbol') return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'object' && v !== null && 'val' in (v as object)) return String((v as { val: unknown }).val) // rich numbers
  return String(v)
}

class ExtraTextPlugin extends FunctionPlugin {
  static implementedFunctions = {
    CONCAT: { method: 'concat', parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1 },
    TEXTJOIN: {
      method: 'textjoin',
      parameters: [
        { argumentType: FunctionArgumentType.STRING },
        { argumentType: FunctionArgumentType.BOOLEAN },
        { argumentType: FunctionArgumentType.ANY }
      ],
      repeatLastArgs: 1
    }
  }
  concat(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('CONCAT'), (...args: unknown[]) => {
      const vals = flatten(args)
      const err = vals.find((v) => v instanceof CellError)
      if (err) return err as CellError
      return vals.map(toText).join('')
    })
  }
  textjoin(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TEXTJOIN'), (delim: string, ignoreEmpty: boolean, ...rest: unknown[]) => {
      const vals = flatten(rest)
      const err = vals.find((v) => v instanceof CellError)
      if (err) return err as CellError
      const parts = vals.map(toText).filter((s) => !(ignoreEmpty && s === ''))
      const out = parts.join(delim)
      return out.length > 32767 ? new CellError(ErrorType.VALUE) : out
    })
  }
}

// ---------- SUBTOTAL / AGGREGATE that honour filtered and hidden rows ----------
type Hidden = (hfSheet: number, row: number) => 0 | 1 | 2 // 0 visible, 1 filtered out, 2 manually hidden
let hiddenRow: Hidden = () => 0
export function setRowVisibility(fn: Hidden): void { hiddenRow = fn }

function num(v: unknown): number | null {
  if (typeof v === 'number') return v
  if (v && typeof v === 'object' && 'val' in (v as object) && typeof (v as { val: unknown }).val === 'number') return (v as { val: number }).val
  return null
}
function stat(fn: number, vals: unknown[]): number | CellError {
  const nums = vals.map(num).filter((x): x is number => x !== null)
  const n = nums.length
  const sum = nums.reduce((a, b) => a + b, 0)
  const mean = n ? sum / n : 0
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0)
  switch (fn) {
    case 1: return n ? mean : new CellError(ErrorType.DIV_BY_ZERO)
    case 2: return n
    case 3: return vals.filter((v) => v !== null && v !== undefined && typeof v !== 'symbol' && v !== '').length
    case 4: return n ? Math.max(...nums) : 0
    case 5: return n ? Math.min(...nums) : 0
    case 6: return nums.reduce((a, b) => a * b, 1)
    case 7: return n > 1 ? Math.sqrt(ss / (n - 1)) : new CellError(ErrorType.DIV_BY_ZERO)
    case 8: return n ? Math.sqrt(ss / n) : new CellError(ErrorType.DIV_BY_ZERO)
    case 9: return sum
    case 10: return n > 1 ? ss / (n - 1) : new CellError(ErrorType.DIV_BY_ZERO)
    case 11: return n ? ss / n : new CellError(ErrorType.DIV_BY_ZERO)
    case 12: { if (!n) return new CellError(ErrorType.NUM); const s = [...nums].sort((a, b) => a - b); return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2 }
    default: return new CellError(ErrorType.VALUE)
  }
}
/** Values of the range args, skipping rows the filter hid (and manually hidden rows when asked). */
function visibleValues(args: unknown[], skipManual: boolean, skipFiltered = true): unknown[] {
  const out: unknown[] = []
  for (const a of args) {
    if (a instanceof SimpleRangeValue) {
      const start = (a.range as any)?.start as { sheet: number; row: number } | undefined
      a.data.forEach((row, i) => {
        if (start) {
          const h = hiddenRow(start.sheet, start.row + i)
          if ((h === 1 && skipFiltered) || (h === 2 && skipManual)) return
        }
        for (const v of row) out.push(v)
      })
    } else out.push(a)
  }
  return out
}

class SubtotalPlugin extends FunctionPlugin {
  static implementedFunctions = {
    SUBTOTAL: { method: 'subtotal', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }], repeatLastArgs: 1, isVolatile: true },
    AGGREGATE: {
      method: 'aggregate',
      parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true }],
      isVolatile: true
    }
  }
  subtotal(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('SUBTOTAL'), (fnNum: number, ...ranges: unknown[]) => {
      const f = Math.trunc(fnNum)
      const base = f > 100 ? f - 100 : f
      if (base < 1 || base > 11) return new CellError(ErrorType.VALUE)
      const vals = visibleValues(ranges, f > 100)
      const err = vals.find((v) => v instanceof CellError)
      return err ? (err as CellError) : stat(base, vals)
    })
  }
  aggregate(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('AGGREGATE'), (fnNum: number, opt: number, range: unknown, k?: number) => {
      const f = Math.trunc(fnNum), o = Math.trunc(opt)
      const ignoreHidden = o === 1 || o === 3 || o === 5 || o === 7
      const ignoreErrors = o === 2 || o === 3 || o === 6 || o === 7
      let vals = visibleValues([range], ignoreHidden, true)
      if (ignoreErrors) vals = vals.filter((v) => !(v instanceof CellError))
      else { const err = vals.find((v) => v instanceof CellError); if (err) return err as CellError }
      if (f === 14 || f === 15) {
        const nums = vals.map(num).filter((x): x is number => x !== null).sort((a, b) => (f === 14 ? b - a : a - b))
        const kk = Math.trunc(k ?? 0)
        return kk >= 1 && kk <= nums.length ? nums[kk - 1] : new CellError(ErrorType.NUM)
      }
      return stat(f, vals)
    })
  }
}

let registered = false
export function registerExtraFunctions(): void {
  if (registered) return
  registered = true
  try { HyperFormula.unregisterFunction('SUBTOTAL') } catch { /* not registered */ }
  HyperFormula.registerFunctionPlugin(SubtotalPlugin, { enGB: { SUBTOTAL: 'SUBTOTAL', AGGREGATE: 'AGGREGATE' }, enUS: { SUBTOTAL: 'SUBTOTAL', AGGREGATE: 'AGGREGATE' } } as any)
  HyperFormula.registerFunctionPlugin(ExtraTextPlugin, { enGB: { CONCAT: 'CONCAT', TEXTJOIN: 'TEXTJOIN' }, enUS: { CONCAT: 'CONCAT', TEXTJOIN: 'TEXTJOIN' } } as any)
}

/** Excel treats bare TRUE/FALSE as literals; HyperFormula only knows TRUE()/FALSE(). */
export const BOOLEAN_NAMES = [
  { name: 'TRUE', expression: '=TRUE()' },
  { name: 'FALSE', expression: '=FALSE()' },
  { name: 'ZZROWVISVER', expression: '=0' } // bumped when rows are filtered/hidden, to recalc SUBTOTAL
]

// Excel functions HyperFormula lacks. Registered once, before any HyperFormula instance is built.
import { formatValue } from '../format/numfmt'
import { registerFastAggregates } from './fastAgg'
import { registerMoreFunctions } from './moreFunctions'
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
/** True when the formula (anywhere in it) calls SUBTOTAL or AGGREGATE. */
function callsSubtotal(ast: any, depth = 0): boolean {
  if (!ast || depth > 50) return false
  if (ast.type === 'FUNCTION_CALL' && (ast.procedureName === 'SUBTOTAL' || ast.procedureName === 'AGGREGATE')) return true
  for (const k of ['args', 'left', 'right', 'value']) {
    const v = ast[k]
    if (Array.isArray(v)) { for (const x of v) if (callsSubtotal(x, depth + 1)) return true }
    else if (v && typeof v === 'object' && v.type && callsSubtotal(v, depth + 1)) return true
  }
  return false
}

/**
 * Values of the range args, skipping rows the filter hid (and manually hidden rows when asked).
 * With `graph`, cells that are themselves SUBTOTAL/AGGREGATE formulas are skipped, as Excel does
 * so a grand total doesn't double-count the group subtotals inside its range.
 */
function visibleValues(args: unknown[], skipManual: boolean, skipFiltered = true, graph?: any): unknown[] {
  const out: unknown[] = []
  for (const a of args) {
    if (a instanceof SimpleRangeValue) {
      const start = (a.range as any)?.start as { sheet: number; row: number; col: number } | undefined
      a.data.forEach((row, i) => {
        if (start) {
          const h = hiddenRow(start.sheet, start.row + i)
          if ((h === 1 && skipFiltered) || (h === 2 && skipManual)) return
        }
        row.forEach((v, j) => {
          if (graph && start && v !== null && v !== undefined) {
            const vx = graph.addressMapping?.getCell?.({ sheet: start.sheet, row: start.row + i, col: start.col + j })
            if (vx && vx.formula && callsSubtotal(vx.formula)) return
          }
          out.push(v)
        })
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
      const vals = visibleValues(ranges, f > 100, true, (this as any).dependencyGraph)
      const err = vals.find((v) => v instanceof CellError)
      return err ? (err as CellError) : stat(base, vals)
    })
  }
  aggregate(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('AGGREGATE'), (fnNum: number, opt: number, range: unknown, k?: number) => {
      const f = Math.trunc(fnNum), o = Math.trunc(opt)
      const ignoreHidden = o === 1 || o === 3 || o === 5 || o === 7
      const ignoreErrors = o === 2 || o === 3 || o === 6 || o === 7
      // Options 0-3 ignore nested SUBTOTAL/AGGREGATE; 4-7 include them.
      let vals = visibleValues([range], ignoreHidden, true, o <= 3 ? (this as any).dependencyGraph : undefined)
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

// ---------- Functions HyperFormula gets wrong or lacks (checked against Excel/LibreOffice) ----------
const DAY_MS = 86400000
function serialToYmd(n: number): { y: number; m: number; d: number } {
  const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * DAY_MS)
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }
}
const daysInMonth = (y: number, m: number): number => new Date(Date.UTC(y, m, 0)).getUTCDate()
function nums(args: unknown[]): number[] {
  return flatten(args).map(num).filter((x): x is number => x !== null)
}
function scalarText(v: unknown): string { return toText(v) }

class ExcelFixPlugin extends FunctionPlugin {
  static implementedFunctions = {
    INT: { method: 'int', parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
    MOD: { method: 'mod', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }] },
    TEXT: { method: 'text', parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.STRING }] },
    DATEDIF: { method: 'datedif', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.STRING }] },
    MODE: { method: 'mode', parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1 },
    'MODE.SNGL': { method: 'mode', parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1 },
    RANK: { method: 'rank', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 }] },
    'RANK.EQ': { method: 'rank', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 }] },
    'RANK.AVG': { method: 'rankavg', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 }] },
    LOOKUP: { method: 'lookup', parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE, optionalArg: true }] },
    FIXED: { method: 'fixed', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 2 }, { argumentType: FunctionArgumentType.BOOLEAN, optionalArg: true, defaultValue: false }] },
    TEXTBEFORE: { method: 'textbefore', parameters: [{ argumentType: FunctionArgumentType.STRING }, { argumentType: FunctionArgumentType.STRING }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 1 }] },
    TEXTAFTER: { method: 'textafter', parameters: [{ argumentType: FunctionArgumentType.STRING }, { argumentType: FunctionArgumentType.STRING }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 1 }] }
  }
  int(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('INT'), (n: number) => Math.floor(n)) // rounds down: INT(-3.5) = -4
  }
  mod(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MOD'), (n: number, d: number) => {
      if (d === 0) return new CellError(ErrorType.DIV_BY_ZERO)
      const r = n - d * Math.floor(n / d) // result takes the divisor's sign: MOD(-7,3) = 2
      return Math.abs(r) < 1e-12 ? 0 : r
    })
  }
  text(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TEXT'), (v: unknown, fmt: string) => {
      if (v instanceof CellError) return v
      let val: unknown = num(v) ?? v
      if (typeof val === 'string' && val.trim() !== '' && !isNaN(Number(val))) val = Number(val)
      if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE'
      if (val === null || val === undefined || typeof val === 'symbol') val = 0
      return formatValue(val as number | string, fmt).text
    })
  }
  datedif(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('DATEDIF'), (s: number, e: number, unit: string) => {
      s = Math.floor(s); e = Math.floor(e)
      if (s > e) return new CellError(ErrorType.NUM)
      const a = serialToYmd(s), b = serialToYmd(e)
      const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0)
      switch (unit.toUpperCase()) {
        case 'Y': return Math.floor(months / 12)
        case 'M': return months
        case 'D': return e - s
        case 'YM': return months % 12
        case 'MD': {
          if (b.d >= a.d) return b.d - a.d
          const pm = b.m === 1 ? 12 : b.m - 1, py = b.m === 1 ? b.y - 1 : b.y
          return daysInMonth(py, pm) - a.d + b.d
        }
        case 'YD': {
          let y = b.y
          const anniv = (yy: number) => Date.UTC(yy, a.m - 1, Math.min(a.d, daysInMonth(yy, a.m))) / DAY_MS
          const end = Date.UTC(b.y, b.m - 1, b.d) / DAY_MS
          if (anniv(y) > end) y--
          return Math.round(end - anniv(y))
        }
        default: return new CellError(ErrorType.NUM)
      }
    })
  }
  mode(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MODE'), (...args: unknown[]) => {
      const list = nums(args)
      const count = new Map<number, number>()
      let best: number | null = null, bestN = 1
      for (const x of list) {
        const n = (count.get(x) ?? 0) + 1
        count.set(x, n)
        if (n > bestN) { bestN = n; best = x }
      }
      // ties: Excel returns the value whose repeat appears first in the data
      if (best !== null) for (const x of list) if (count.get(x) === bestN) return x
      return new CellError(ErrorType.NA)
    })
  }
  private rankOf(n: number, ref: unknown, order: number, avg: boolean): number | CellError {
    const list = nums([ref])
    if (!list.includes(n)) return new CellError(ErrorType.NA)
    const better = list.filter((x) => (order ? x < n : x > n)).length
    const ties = list.filter((x) => x === n).length
    return avg ? better + (ties + 1) / 2 : better + 1
  }
  rank(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('RANK'), (n: number, ref: unknown, order: number) => this.rankOf(n, ref, order, false))
  }
  rankavg(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('RANK.AVG'), (n: number, ref: unknown, order: number) => this.rankOf(n, ref, order, true))
  }
  lookup(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LOOKUP'), (v: unknown, look: SimpleRangeValue, res?: SimpleRangeValue) => {
      const grid = look.data
      const h = grid.length, w = grid[0]?.length ?? 0
      let keys: unknown[], vals: unknown[]
      if (res) {
        keys = grid.flat()
        vals = res.data.flat()
      } else if (w > h) { keys = grid[0]; vals = grid[h - 1] } else { keys = grid.map((r) => r[0]); vals = grid.map((r) => r[w - 1]) }
      // Approximate match on sorted data: the last key <= v of the same kind (number vs text).
      const vn = num(v)
      let hit = -1
      keys.forEach((k, i) => {
        if (vn !== null) { const kn = num(k); if (kn !== null && kn <= vn) hit = i }
        else if (typeof k === 'string' && typeof v === 'string' && k.localeCompare(v, undefined, { sensitivity: 'base' }) <= 0) hit = i
      })
      if (hit < 0 || hit >= vals.length) return new CellError(ErrorType.NA)
      const out = vals[hit]
      return out === undefined || typeof out === 'symbol' ? 0 : (out as InterpreterValue)
    })
  }
  fixed(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('FIXED'), (n: number, dec: number, noCommas: boolean) => {
      const d = Math.trunc(dec)
      const rounded = d >= 0 ? Number(n.toFixed(Math.min(d, 100))) : Math.round(n / 10 ** -d) * 10 ** -d
      return rounded.toLocaleString('en-US', { minimumFractionDigits: Math.max(0, d), maximumFractionDigits: Math.max(0, d), useGrouping: !noCommas })
    })
  }
  private around(txt: string, delim: string, inst: number, before: boolean): string | CellError {
    if (inst === 0 || delim === '') return new CellError(ErrorType.VALUE)
    let idx = -1
    if (inst > 0) { let from = 0; for (let k = 0; k < inst; k++) { idx = txt.indexOf(delim, from); if (idx < 0) break; from = idx + delim.length } }
    else { let from = txt.length; for (let k = 0; k < -inst; k++) { idx = txt.lastIndexOf(delim, from); if (idx < 0) break; from = idx - 1 } }
    if (idx < 0) return new CellError(ErrorType.NA)
    return before ? txt.slice(0, idx) : txt.slice(idx + delim.length)
  }
  textbefore(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TEXTBEFORE'), (t: string, d: string, i: number) => this.around(t, d, Math.trunc(i), true))
  }
  textafter(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TEXTAFTER'), (t: string, d: string, i: number) => this.around(t, d, Math.trunc(i), false))
  }
}
void scalarText

// ---------- INDIRECT: a reference built from text (values only; volatile, so it follows every edit) ----------
const colIdx = (t: string): number => [...t.toUpperCase()].reduce((v, ch) => v * 26 + ch.charCodeAt(0) - 64, 0) - 1
class IndirectPlugin extends FunctionPlugin {
  static implementedFunctions = {
    INDIRECT: { method: 'indirect', parameters: [{ argumentType: FunctionArgumentType.STRING }, { argumentType: FunctionArgumentType.BOOLEAN, optionalArg: true, defaultValue: true }], isVolatile: true }
  }
  indirect(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('INDIRECT'), (text: string, a1: boolean) => {
      if (!a1) return new CellError(ErrorType.REF) // R1C1 text is not supported
      const graph = (this as any).dependencyGraph
      let ref = text.trim()
      let sheet: number = state.formulaAddress.sheet
      const bang = ref.lastIndexOf('!')
      if (bang >= 0) {
        let name = ref.slice(0, bang)
        if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'")
        const id = graph.sheetMapping.getSheetId(name)
        if (id === undefined) return new CellError(ErrorType.REF)
        sheet = id
        ref = ref.slice(bang + 1)
      }
      const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref)
      if (!m) return new CellError(ErrorType.REF)
      const c1 = colIdx(m[1]), r1 = +m[2] - 1
      const c2 = m[3] ? colIdx(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1
      const [ra, rb, ca, cb] = [Math.min(r1, r2), Math.max(r1, r2), Math.min(c1, c2), Math.max(c1, c2)]
      if (ra < 0 || (rb - ra + 1) * (cb - ca + 1) > 1_000_000) return new CellError(ErrorType.REF)
      const at = (r: number, c: number): unknown => graph.getScalarValue({ sheet, row: r, col: c })
      if (ra === rb && ca === cb) return at(ra, ca) as never
      const data: unknown[][] = []
      for (let r = ra; r <= rb; r++) { const row: unknown[] = []; for (let c = ca; c <= cb; c++) row.push(at(r, c)); data.push(row) }
      return SimpleRangeValue.onlyValues(data as never) as never
    })
  }
}

let registered = false
export function registerExtraFunctions(): void {
  if (registered) return
  registered = true
  for (const f of ['SUBTOTAL', 'INT', 'MOD', 'TEXT', 'DATEDIF']) {
    try { HyperFormula.unregisterFunction(f) } catch { /* not registered */ }
  }
  const fixNames = Object.fromEntries(Object.keys(ExcelFixPlugin.implementedFunctions).map((k) => [k, k]))
  HyperFormula.registerFunctionPlugin(ExcelFixPlugin, { enGB: fixNames, enUS: fixNames } as any)
  HyperFormula.registerFunctionPlugin(SubtotalPlugin, { enGB: { SUBTOTAL: 'SUBTOTAL', AGGREGATE: 'AGGREGATE' }, enUS: { SUBTOTAL: 'SUBTOTAL', AGGREGATE: 'AGGREGATE' } } as any)
  registerFastAggregates()
  registerMoreFunctions()
  HyperFormula.registerFunctionPlugin(IndirectPlugin, { enGB: { INDIRECT: 'INDIRECT' }, enUS: { INDIRECT: 'INDIRECT' } } as any)
  HyperFormula.registerFunctionPlugin(ExtraTextPlugin, { enGB: { CONCAT: 'CONCAT', TEXTJOIN: 'TEXTJOIN' }, enUS: { CONCAT: 'CONCAT', TEXTJOIN: 'TEXTJOIN' } } as any)
}

/** Excel treats bare TRUE/FALSE as literals; HyperFormula only knows TRUE()/FALSE(). */
export const BOOLEAN_NAMES = [
  { name: 'TRUE', expression: '=TRUE()' },
  { name: 'FALSE', expression: '=FALSE()' },
  { name: 'ZZROWVISVER', expression: '=0' } // bumped when rows are filtered/hidden, to recalc SUBTOTAL
]

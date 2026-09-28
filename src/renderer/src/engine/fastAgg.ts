// Fast COUNTIF(S) / SUMIF(S) / AVERAGEIF(S) / MAXIFS / MINIFS / XLOOKUP / INDEX.
//
// HyperFormula evaluates these by re-reading and re-parsing every cell of the ranges for every formula (about 7 ms per
// call on 5,000 rows, so a column of 5,000 SUMIFS takes 35 s). Here each range is flattened once per "epoch" (any
// engine mutation starts a new one) and equality criteria go through a hash index, so a call costs O(matches).
import { FunctionPlugin, FunctionArgumentType, HyperFormula, SimpleRangeValue, CellError, ErrorType, ArraySize } from 'hyperformula'
type InterpreterValue = any

// ---------- epochs and caches ----------
let epoch = 0
/** Called before every engine mutation: flattened ranges and indexes built earlier may be stale. */
export function bumpAggEpoch(): void { epoch++ }

interface Line { vals: unknown[]; index?: Map<string, number[]> }
interface Flat { vals: unknown[]; h: number; w: number; index?: Map<string, number[]>; lines?: Map<string, Line> }
const caches = new WeakMap<object, { epoch: number; map: Map<string, Flat> }>()

const isBlank = (v: unknown): boolean => v === null || v === undefined || typeof v === 'symbol' || v === ''
function norm(v: unknown): unknown {
  if (typeof v === 'symbol' || v === undefined) return null
  if (v && typeof v === 'object' && !(v instanceof CellError) && 'val' in (v as object)) return (v as { val: unknown }).val
  return v
}

/** Row-major values of a range argument, cached while nothing has been edited. */
function flatten(plugin: any, arg: any): Flat {
  const range = arg?.range
  const graph = plugin.dependencyGraph
  if (arg instanceof SimpleRangeValue && range && graph) {
    let c = caches.get(graph)
    if (!c || c.epoch !== epoch) { c = { epoch, map: new Map() }; caches.set(graph, c) }
    const h = arg.height(), w = arg.width()
    const key = `${range.start.sheet}:${range.start.row}:${range.start.col}:${h}:${w}`
    let f = c.map.get(key)
    if (!f) {
      const vals: unknown[] = new Array(h * w)
      let k = 0
      for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) vals[k++] = norm(graph.getScalarValue({ sheet: range.start.sheet, row: range.start.row + i, col: range.start.col + j }))
      f = { vals, h, w }
      c.map.set(key, f)
    }
    return f
  }
  if (arg instanceof SimpleRangeValue) {
    const data = arg.data as unknown[][]
    return { vals: data.flatMap((row) => row.map(norm)), h: data.length, w: data[0]?.length ?? 0 }
  }
  return { vals: [norm(arg)], h: 1, w: 1 }
}

// ---------- criteria ----------
const numberLike = (s: string): number | null => {
  const t = s.trim()
  if (t === '' || !/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?%?$/i.test(t)) return null
  const n = parseFloat(t)
  return t.endsWith('%') ? n / 100 : n
}
function wildcardRe(pat: string): RegExp {
  let out = ''
  for (let i = 0; i < pat.length; i++) {
    const ch = pat[i]
    if (ch === '~' && (pat[i + 1] === '*' || pat[i + 1] === '?' || pat[i + 1] === '~')) { out += '\\' + pat[++i]; continue }
    out += ch === '*' ? '[\\s\\S]*' : ch === '?' ? '[\\s\\S]' : ch.replace(/[.+^${}()|[\]\\]/g, '\\$&')
  }
  return new RegExp('^' + out + '$', 'i')
}
const hasWildcard = (s: string): boolean => /[*?]/.test(s)

interface Compiled { eqKeys?: string[]; test: (v: unknown) => boolean }
const keyOf = (v: unknown): string | null => (typeof v === 'number' ? 'n:' + v : typeof v === 'string' ? 's:' + v.toLowerCase() : typeof v === 'boolean' ? 'b:' + v : null)

export function compileCriterion(crit: unknown): Compiled {
  crit = norm(crit)
  if (crit instanceof CellError) return { test: () => false }
  if (typeof crit === 'number') return { eqKeys: ['n:' + crit, 's:' + String(crit)], test: (v) => v === crit || (typeof v === 'string' && numberLike(v) === crit) }
  if (typeof crit === 'boolean') return { eqKeys: ['b:' + crit], test: (v) => v === crit }
  if (crit === null || crit === undefined) return { test: (v) => isBlank(v) || v === 0 } // an empty criterion cell matches like 0? Excel: like =0
  const s = String(crit)
  const m = /^(>=|<=|<>|>|<|=)?([\s\S]*)$/.exec(s)!
  const op = m[1] ?? ''
  const operand = m[2]
  if (operand === '') {
    if (op === '=') return { test: (v) => v === null } // "=" matches only truly empty cells, "" also matches empty strings
    if (op === '<>') return { test: (v) => v !== null }
    if (op === '') return { test: isBlank }
    return { test: () => false }
  }
  const n = numberLike(operand)
  if (n !== null) {
    switch (op) {
      case '>': return { test: (v) => typeof v === 'number' && v > n }
      case '<': return { test: (v) => typeof v === 'number' && v < n }
      case '>=': return { test: (v) => typeof v === 'number' && v >= n }
      case '<=': return { test: (v) => typeof v === 'number' && v <= n }
      case '<>': return { test: (v) => v !== n }
      default: return { eqKeys: ['n:' + n, 's:' + operand.toLowerCase()], test: (v) => v === n || (typeof v === 'string' && (numberLike(v) === n || v.toLowerCase() === operand.toLowerCase())) }
    }
  }
  const up = operand.toUpperCase()
  if (up === 'TRUE' || up === 'FALSE') {
    const b = up === 'TRUE'
    if (op === '' || op === '=') return { eqKeys: ['b:' + b, 's:' + operand.toLowerCase()], test: (v) => v === b || (typeof v === 'string' && v.toLowerCase() === operand.toLowerCase()) }
    if (op === '<>') return { test: (v) => v !== b }
  }
  if (hasWildcard(operand) && (op === '' || op === '=' || op === '<>')) {
    const re = wildcardRe(operand)
    return op === '<>' ? { test: (v) => !(typeof v === 'string' && re.test(v)) } : { test: (v) => typeof v === 'string' && re.test(v) }
  }
  const lo = operand.toLowerCase()
  switch (op) {
    case '': case '=': return { eqKeys: ['s:' + lo], test: (v) => typeof v === 'string' && v.toLowerCase() === lo }
    case '<>': return { test: (v) => !(typeof v === 'string' && v.toLowerCase() === lo) }
    case '>': return { test: (v) => typeof v === 'string' && v.toLowerCase().localeCompare(lo) > 0 }
    case '<': return { test: (v) => typeof v === 'string' && v.toLowerCase().localeCompare(lo) < 0 }
    case '>=': return { test: (v) => typeof v === 'string' && v.toLowerCase().localeCompare(lo) >= 0 }
    default: return { test: (v) => typeof v === 'string' && v.toLowerCase().localeCompare(lo) <= 0 }
  }
}

function indexOf(f: Flat): Map<string, number[]> {
  if (f.index) return f.index
  const m = new Map<string, number[]>()
  for (let i = 0; i < f.vals.length; i++) {
    const k = keyOf(f.vals[i])
    if (k === null) continue
    const a = m.get(k)
    if (a) a.push(i)
    else m.set(k, [i])
  }
  f.index = m
  return m
}

/** Row-major positions satisfying every (range, criterion) pair. */
function matching(pairs: { f: Flat; c: Compiled }[]): number[] {
  // Start from the most selective equality criterion, else scan the first range.
  let base: number[] | null = null
  let baseIdx = -1
  for (let p = 0; p < pairs.length; p++) {
    const { f, c } = pairs[p]
    if (!c.eqKeys) continue
    const idx = indexOf(f)
    const parts = c.eqKeys.map((k) => idx.get(k)).filter((x): x is number[] => !!x)
    const cand = parts.length === 1 ? parts[0] : parts.length === 0 ? [] : parts.flat().sort((a, b) => a - b)
    if (base === null || cand.length < base.length) { base = cand; baseIdx = p }
  }
  if (base === null) {
    const { f, c } = pairs[0]
    base = []
    for (let i = 0; i < f.vals.length; i++) if (c.test(f.vals[i])) base.push(i)
    baseIdx = 0
  }
  let out = base
  for (let p = 0; p < pairs.length; p++) {
    if (p === baseIdx) continue
    const { f, c } = pairs[p]
    out = out.filter((i) => c.test(f.vals[i]))
    if (!out.length) break
  }
  return out
}

const VALUE = (): CellError => new CellError(ErrorType.VALUE)

class FastAggPlugin extends FunctionPlugin {
  static implementedFunctions = {
    COUNTIF: { method: 'countif', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }] },
    COUNTIFS: { method: 'countifs', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 2 },
    SUMIF: { method: 'sumif', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.RANGE, optionalArg: true }] },
    SUMIFS: { method: 'sumifs', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 2 },
    AVERAGEIF: { method: 'averageif', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.RANGE, optionalArg: true }] },
    AVERAGEIFS: { method: 'averageifs', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 2 },
    MAXIFS: { method: 'maxifs', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 2 },
    MINIFS: { method: 'minifs', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 2 }
  }

  private pairs(rest: unknown[], first?: { range: unknown; crit: unknown }): { f: Flat; c: Compiled }[] | CellError {
    const out: { f: Flat; c: Compiled }[] = []
    const list = first ? [first.range, first.crit, ...rest] : rest
    for (let i = 0; i < list.length; i += 2) {
      const f = flatten(this, list[i])
      out.push({ f, c: compileCriterion(list[i + 1]) })
    }
    const { h, w } = out[0].f
    if (out.some((p) => p.f.h !== h || p.f.w !== w)) return VALUE()
    return out
  }
  private run(ast: any, state: any, name: string, body: (...a: any[]) => unknown): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata(name), body as never)
  }
  private numsAt(f: Flat, idx: number[]): number[] {
    const out: number[] = []
    for (const i of idx) { const v = f.vals[i]; if (typeof v === 'number') out.push(v) }
    return out
  }

  countif(ast: any, state: any): InterpreterValue {
    return this.run(ast, state, 'COUNTIF', (range: unknown, crit: unknown) => matching([{ f: flatten(this, range), c: compileCriterion(crit) }]).length)
  }
  countifs(ast: any, state: any): InterpreterValue {
    return this.run(ast, state, 'COUNTIFS', (...args: unknown[]) => {
      const p = this.pairs(args)
      return p instanceof CellError ? p : matching(p).length
    })
  }
  sumif(ast: any, state: any): InterpreterValue {
    return this.run(ast, state, 'SUMIF', (range: unknown, crit: unknown, sum?: unknown) => {
      const f = flatten(this, range)
      const idx = matching([{ f, c: compileCriterion(crit) }])
      const sf = sum === undefined ? f : flatten(this, sum)
      return this.numsAt(sf, idx).reduce((a, b) => a + b, 0)
    })
  }
  sumifs(ast: any, state: any): InterpreterValue {
    return this.run(ast, state, 'SUMIFS', (sum: unknown, ...args: unknown[]) => {
      const sf = flatten(this, sum)
      const p = this.pairs(args)
      if (p instanceof CellError) return p
      if (p[0].f.h !== sf.h || p[0].f.w !== sf.w) return VALUE()
      return this.numsAt(sf, matching(p)).reduce((a, b) => a + b, 0)
    })
  }
  averageif(ast: any, state: any): InterpreterValue {
    return this.run(ast, state, 'AVERAGEIF', (range: unknown, crit: unknown, avg?: unknown) => {
      const f = flatten(this, range)
      const idx = matching([{ f, c: compileCriterion(crit) }])
      const n = this.numsAt(avg === undefined ? f : flatten(this, avg), idx)
      return n.length ? n.reduce((a, b) => a + b, 0) / n.length : new CellError(ErrorType.DIV_BY_ZERO)
    })
  }
  averageifs(ast: any, state: any): InterpreterValue {
    return this.run(ast, state, 'AVERAGEIFS', (avg: unknown, ...args: unknown[]) => {
      const sf = flatten(this, avg)
      const p = this.pairs(args)
      if (p instanceof CellError) return p
      if (p[0].f.h !== sf.h || p[0].f.w !== sf.w) return VALUE()
      const n = this.numsAt(sf, matching(p))
      return n.length ? n.reduce((a, b) => a + b, 0) / n.length : new CellError(ErrorType.DIV_BY_ZERO)
    })
  }
  private extreme(ast: any, state: any, name: 'MAXIFS' | 'MINIFS'): InterpreterValue {
    return this.run(ast, state, name, (target: unknown, ...args: unknown[]) => {
      const sf = flatten(this, target)
      const p = this.pairs(args)
      if (p instanceof CellError) return p
      if (p[0].f.h !== sf.h || p[0].f.w !== sf.w) return VALUE()
      const n = this.numsAt(sf, matching(p))
      if (!n.length) return 0
      return name === 'MAXIFS' ? n.reduce((a, b) => (b > a ? b : a)) : n.reduce((a, b) => (b < a ? b : a))
    })
  }
  maxifs(ast: any, state: any): InterpreterValue { return this.extreme(ast, state, 'MAXIFS') }
  minifs(ast: any, state: any): InterpreterValue { return this.extreme(ast, state, 'MINIFS') }
}

// ---------- XLOOKUP / INDEX ----------
class FastLookupPlugin extends FunctionPlugin {
  static implementedFunctions = {
    XLOOKUP: {
      method: 'xlookup',
      sizeOfResultArrayMethod: 'xlookupArraySize',
      parameters: [
        { argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.ANY, optionalArg: true }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 1 }
      ]
    },
    INDEX: {
      method: 'index',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 }]
    }
  }
  /** INDEX(array, row, [col]) without copying the array for every call. */
  index(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('INDEX'), ((arr: any, row: number, col: number) => {
      const h = arr.height(), w = arr.width()
      let r = Math.trunc(row), c = Math.trunc(col)
      // One-row or one-column arrays take a single index as their position.
      if (h === 1 && w > 1 && col === 0 && ast.args.length < 3) { c = r; r = 1 } else if (w === 1 && h > 1 && col === 0) { c = 1 }
      if (r < 0 || c < 0 || r > h || c > w) return new CellError(ErrorType.REF)
      const graph = (this as any).dependencyGraph
      const at = (i: number, j: number): unknown => {
        if (arr.range && graph) return graph.getScalarValue({ sheet: arr.range.start.sheet, row: arr.range.start.row + i, col: arr.range.start.col + j })
        return arr.data[i][j]
      }
      if (r > 0 && c > 0) { const v = at(r - 1, c - 1); return (typeof v === 'symbol' ? 0 : v) as never }
      if (r === 0 && c === 0) return arr as never
      if (r === 0) return SimpleRangeValue.onlyValues(Array.from({ length: h }, (_, i) => [at(i, c - 1)]) as never) as never
      return SimpleRangeValue.onlyValues([Array.from({ length: w }, (_, j) => at(r - 1, j))] as never) as never
    }) as never)
  }
  xlookupArraySize(ast: any): ArraySize {
    const l = ast?.args?.[1], r = ast?.args?.[2]
    if (l?.start == null || l?.end == null || r?.start == null || r?.end == null) return ArraySize.error()
    const lh = l.end.row - l.start.row + 1, lw = l.end.col - l.start.col + 1
    const rh = r.end.row - r.start.row + 1, rw = r.end.col - r.start.col + 1
    if (lw === 1 && rh === lh) return new ArraySize(rw, 1)
    if (lh === 1 && rw === lw) return new ArraySize(1, rh)
    return ArraySize.error()
  }
  xlookup(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('XLOOKUP'), ((key: unknown, look: unknown, ret: unknown, ifNot: unknown, mode: number, search: number) => {
      const lf = flatten(this, look)
      const rf = flatten(this, ret)
      const vertical = lf.w === 1
      if (lf.h !== 1 && lf.w !== 1) return VALUE()
      if ((vertical && rf.h !== lf.h) || (!vertical && rf.w !== lf.w)) return VALUE()
      const k = norm(key)
      if (k instanceof CellError) return k
      let at = -1
      const mm = Math.trunc(mode), sm = Math.trunc(search)
      if (mm === 0 || mm === 2) {
        const c = mm === 2 && typeof k === 'string' ? { test: (v: unknown) => typeof v === 'string' && wildcardRe(k).test(v) } : compileCriterion(typeof k === 'string' ? '=' + escapeEq(k) : k)
        const eq = mm === 0 ? keyOf(k) : null
        if (eq !== null) {
          const list = indexOf(lf).get(eq)
          if (list && list.length) at = sm === -1 || sm === -2 ? list[list.length - 1] : list[0]
        } else {
          const rev = sm === -1 || sm === -2
          for (let i = rev ? lf.vals.length - 1 : 0; rev ? i >= 0 : i < lf.vals.length; rev ? i-- : i++) if (c.test(lf.vals[i])) { at = i; break }
        }
      } else {
        // -1: exact or next smaller; 1: exact or next larger (numbers and text compared like Excel)
        let best = -1
        for (let i = 0; i < lf.vals.length; i++) {
          const v = lf.vals[i]
          if (isBlank(v) || typeof v !== typeof k) continue
          const cmp = typeof v === 'number' ? v - (k as number) : String(v).toLowerCase().localeCompare(String(k).toLowerCase())
          if (cmp === 0) { best = i; break }
          if (mm === -1 ? cmp < 0 : cmp > 0) {
            if (best < 0 || (mm === -1 ? cmp > cmpOf(lf.vals[best], k) : cmp < cmpOf(lf.vals[best], k))) best = i
          }
        }
        at = best
      }
      if (at < 0) return ifNot === undefined || ifNot === null ? new CellError(ErrorType.NA) : ifNot
      if (vertical) {
        if (rf.w === 1) return rf.vals[at] === null ? 0 : (rf.vals[at] as never)
        return SimpleRangeValue.onlyValues([rf.vals.slice(at * rf.w, at * rf.w + rf.w)] as never) as never
      }
      if (rf.h === 1) return rf.vals[at] === null ? 0 : (rf.vals[at] as never)
      return SimpleRangeValue.onlyValues(Array.from({ length: rf.h }, (_, r) => [rf.vals[r * rf.w + at]]) as never) as never
    }) as never)
  }
}
// ---------- VLOOKUP / HLOOKUP / MATCH ----------
/** One column ('c') or row ('r') of a flattened range, with its own hash index. */
function lineOf(f: Flat, dim: 'c' | 'r', k: number): Line {
  const m = (f.lines ??= new Map())
  const key = dim + k
  let l = m.get(key)
  if (!l) {
    l = { vals: dim === 'c' ? Array.from({ length: f.h }, (_, i) => f.vals[i * f.w + k]) : f.vals.slice(k * f.w, (k + 1) * f.w) }
    m.set(key, l)
  }
  return l
}
function lineIndex(l: Line): Map<string, number[]> {
  if (l.index) return l.index
  const m = new Map<string, number[]>()
  for (let i = 0; i < l.vals.length; i++) {
    const k = keyOf(l.vals[i])
    if (k === null) continue
    const a = m.get(k)
    if (a) a.push(i)
    else m.set(k, [i])
  }
  l.index = m
  return m
}
function findExact(l: Line, key: unknown): number {
  if (typeof key === 'string' && hasWildcard(key)) {
    const re = wildcardRe(key)
    return l.vals.findIndex((v) => typeof v === 'string' && re.test(v))
  }
  const k = keyOf(key)
  if (k === null) return -1
  const a = lineIndex(l).get(k)
  return a ? a[0] : -1
}
const sameType = (a: unknown, b: unknown): boolean => typeof a === typeof b && !(a instanceof CellError)
const cmpVals = (a: unknown, b: unknown): number => (typeof a === 'number' ? a - (b as number) : typeof a === 'boolean' ? Number(a) - Number(b) : String(a).toLowerCase().localeCompare(String(b).toLowerCase()))
/** Sorted-data search: index of the last value <= key (ascending) or of the last value >= key (descending); -1 if none. */
function findApprox(vals: unknown[], key: unknown, descending: boolean): number {
  let lo = 0, hi = vals.length - 1, best = -1
  while (lo <= hi) {
    let mid = (lo + hi) >> 1
    // skip cells of another type (blanks, text among numbers) like Excel does
    let m = mid
    while (m <= hi && !sameType(vals[m], key)) m++
    if (m > hi) { m = mid; while (m >= lo && !sameType(vals[m], key)) m-- }
    if (m < lo) break
    mid = m
    const c = cmpVals(vals[mid], key)
    if (c === 0 || (descending ? c > 0 : c < 0)) { best = mid; lo = mid + 1 } else hi = mid - 1
  }
  return best
}
const keyArg = (k: unknown): unknown => (k instanceof SimpleRangeValue ? norm(k.data[0]?.[0]) : norm(k))

class FastVLookupPlugin extends FunctionPlugin {
  static implementedFunctions = {
    VLOOKUP: { method: 'vlookup', parameters: [{ argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.BOOLEAN, optionalArg: true, defaultValue: true }] },
    HLOOKUP: { method: 'hlookup', parameters: [{ argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.BOOLEAN, optionalArg: true, defaultValue: true }] },
    MATCH: { method: 'match', parameters: [{ argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 1 }] }
  }
  private lookup(ast: any, state: any, name: 'VLOOKUP' | 'HLOOKUP'): InterpreterValue {
    const v = name === 'VLOOKUP'
    return this.runFunction(ast.args, state, this.metadata(name), ((key0: unknown, table: unknown, idx: number, approx: boolean) => {
      const key = keyArg(key0)
      if (key instanceof CellError) return key
      const f = flatten(this, table)
      const n = Math.trunc(idx)
      if (n < 1) return VALUE()
      if (n > (v ? f.w : f.h)) return new CellError(ErrorType.REF)
      if (key === null) return new CellError(ErrorType.NA)
      const l = lineOf(f, v ? 'c' : 'r', 0)
      const at = approx ? findApprox(l.vals, key, false) : findExact(l, key)
      if (at < 0) return new CellError(ErrorType.NA)
      const out = v ? f.vals[at * f.w + n - 1] : f.vals[(n - 1) * f.w + at]
      return (out === null || out === undefined ? 0 : out) as never
    }) as never)
  }
  vlookup(ast: any, state: any): InterpreterValue { return this.lookup(ast, state, 'VLOOKUP') }
  hlookup(ast: any, state: any): InterpreterValue { return this.lookup(ast, state, 'HLOOKUP') }
  match(ast: any, state: any): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MATCH'), ((key0: unknown, range: unknown, type: number) => {
      const key = keyArg(key0)
      if (key instanceof CellError) return key
      const f = flatten(this, range)
      if (f.h !== 1 && f.w !== 1) return new CellError(ErrorType.NA)
      if (key === null) return new CellError(ErrorType.NA)
      const l: Line = f.h === 1 ? lineOf(f, 'r', 0) : lineOf(f, 'c', 0)
      const t = Math.sign(Math.trunc(type))
      const at = t === 0 ? findExact(l, key) : findApprox(l.vals, key, t < 0)
      return at < 0 ? new CellError(ErrorType.NA) : at + 1
    }) as never)
  }
}

const escapeEq = (s: string): string => s.replace(/~/g, '~~').replace(/\*/g, '~*').replace(/\?/g, '~?')
const cmpOf = (v: unknown, k: unknown): number => (typeof v === 'number' ? v - (k as number) : String(v).toLowerCase().localeCompare(String(k).toLowerCase()))

let registered = false
export function registerFastAggregates(): void {
  if (registered) return
  registered = true
  const names = (p: { implementedFunctions: object }): Record<string, string> => Object.fromEntries(Object.keys(p.implementedFunctions).map((k) => [k, k]))
  for (const p of [FastAggPlugin, FastLookupPlugin, FastVLookupPlugin]) {
    for (const f of Object.keys(p.implementedFunctions)) { try { HyperFormula.unregisterFunction(f) } catch { /* not registered */ } }
    HyperFormula.registerFunctionPlugin(p as never, { enGB: names(p), enUS: names(p) } as never)
  }
}

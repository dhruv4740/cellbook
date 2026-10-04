// More Excel functions HyperFormula lacks: TAKE, DROP, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW, SORTBY, MUNIT, MINVERSE,
// MDETERM, FREQUENCY, TREND, FORECAST, INTERCEPT, PERCENTRANK, NUMBERVALUE, TYPE, VALUETOTEXT, TEXTSPLIT, WRAPROWS, WRAPCOLS,
// EXPAND, RANDARRAY, ARRAYTOTEXT, CELL, INFO.
// Array functions declare their result size from the formula text, so their row / column counts must be literal numbers.
import { FunctionPlugin, FunctionArgumentType, HyperFormula, SimpleRangeValue, CellError, ErrorType, ArraySize } from 'hyperformula'
type Any = any // eslint-disable-line @typescript-eslint/no-explicit-any

const A = FunctionArgumentType
const ANY = { argumentType: A.ANY }
const RANGE = { argumentType: A.RANGE }
const NUM = { argumentType: A.NUMBER }
const optNum = (d: number) => ({ argumentType: A.NUMBER, optionalArg: true, defaultValue: d })

const dims = (n: Any): [number, number] | null => (n?.start && n?.end ? [n.end.row - n.start.row + 1, n.end.col - n.start.col + 1] : null)
const lit = (n: Any): number | null => {
  if (!n) return null
  if (typeof n.value === 'number') return n.value
  if (n.type === 'MINUS_UNARY_OP') { const v = lit(n.value); return v === null ? null : -v }
  return null
}
const bad = (): ArraySize => ArraySize.error()
const VALUE = (): CellError => new CellError(ErrorType.VALUE)
const NA = (): CellError => new CellError(ErrorType.NA)
const DIV0 = (): CellError => new CellError(ErrorType.DIV_BY_ZERO)

function num(v: unknown): number | null {
  if (typeof v === 'number') return v
  if (v && typeof v === 'object' && 'val' in (v as object) && typeof (v as { val: unknown }).val === 'number') return (v as { val: number }).val
  return null
}
const grid = (r: unknown): unknown[][] => (r instanceof SimpleRangeValue ? r.data : [[r]])
const out = (data: unknown[][]): never => SimpleRangeValue.onlyValues(data as never) as never
const firstError = (rows: unknown[][]): CellError | undefined => { for (const r of rows) for (const v of r) if (v instanceof CellError) return v }

function pairs(ys: unknown[][], xs: unknown[][]): { x: number; y: number }[] | CellError {
  const y = ys.flat(), x = xs.flat()
  if (x.length !== y.length) return NA()
  const out: { x: number; y: number }[] = []
  for (let i = 0; i < x.length; i++) { const a = num(x[i]), b = num(y[i]); if (a !== null && b !== null) out.push({ x: a, y: b }) }
  return out
}
function line(p: { x: number; y: number }[]): { slope: number; intercept: number } | null {
  const n = p.length
  if (n < 2) return null
  const mx = p.reduce((s, q) => s + q.x, 0) / n, my = p.reduce((s, q) => s + q.y, 0) / n
  let sxx = 0, sxy = 0
  for (const q of p) { sxx += (q.x - mx) ** 2; sxy += (q.x - mx) * (q.y - my) }
  if (sxx === 0) return null
  const slope = sxy / sxx
  return { slope, intercept: my - slope * mx }
}

function luInverse(m: number[][]): { inv: number[][]; det: number } | null {
  const n = m.length
  const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))])
  let det = 1
  for (let c = 0; c < n; c++) {
    let p = c
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r
    if (Math.abs(a[p][c]) < 1e-12) return null
    if (p !== c) { [a[p], a[c]] = [a[c], a[p]]; det = -det }
    det *= a[c][c]
    const d = a[c][c]
    for (let k = 0; k < 2 * n; k++) a[c][k] /= d
    for (let r = 0; r < n; r++) if (r !== c) { const f = a[r][c]; if (f) for (let k = 0; k < 2 * n; k++) a[r][k] -= f * a[c][k] }
  }
  return { inv: a.map((r) => r.slice(n)), det }
}
function determinant(m: number[][]): number {
  const n = m.length
  const a = m.map((r) => [...r])
  let det = 1
  for (let c = 0; c < n; c++) {
    let p = c
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r
    if (Math.abs(a[p][c]) < 1e-14) return 0
    if (p !== c) { [a[p], a[c]] = [a[c], a[p]]; det = -det }
    det *= a[c][c]
    for (let r = c + 1; r < n; r++) { const f = a[r][c] / a[c][c]; for (let k = c; k < n; k++) a[r][k] -= f * a[c][k] }
  }
  return det
}
const square = (g: unknown[][]): number[][] | null => {
  const n = g.length
  if (!n || g.some((r) => r.length !== n)) return null
  const m = g.map((r) => r.map((v) => num(v)))
  return m.every((r) => r.every((v) => v !== null)) ? (m as number[][]) : null
}

class MorePlugin extends FunctionPlugin {
  static implementedFunctions = {
    TAKE: { method: 'take', sizeOfResultArrayMethod: 'takeSize', parameters: [RANGE, NUM, { argumentType: A.NUMBER, optionalArg: true }] },
    DROP: { method: 'drop', sizeOfResultArrayMethod: 'dropSize', parameters: [RANGE, NUM, { argumentType: A.NUMBER, optionalArg: true }] },
    CHOOSECOLS: { method: 'choosecols', sizeOfResultArrayMethod: 'choosecolsSize', parameters: [RANGE, NUM], repeatLastArgs: 1 },
    CHOOSEROWS: { method: 'chooserows', sizeOfResultArrayMethod: 'chooserowsSize', parameters: [RANGE, NUM], repeatLastArgs: 1 },
    TOCOL: { method: 'tocol', sizeOfResultArrayMethod: 'tocolSize', parameters: [RANGE, optNum(0), { argumentType: A.BOOLEAN, optionalArg: true, defaultValue: false }] },
    TOROW: { method: 'torow', sizeOfResultArrayMethod: 'torowSize', parameters: [RANGE, optNum(0), { argumentType: A.BOOLEAN, optionalArg: true, defaultValue: false }] },
    SORTBY: { method: 'sortby', sizeOfResultArrayMethod: 'sortbySize', parameters: [RANGE, RANGE, optNum(1)], repeatLastArgs: 2 },
    MUNIT: { method: 'munit', sizeOfResultArrayMethod: 'munitSize', parameters: [NUM] },
    MINVERSE: { method: 'minverse', sizeOfResultArrayMethod: 'sameSize', parameters: [RANGE] },
    MDETERM: { method: 'mdeterm', parameters: [RANGE] },
    FREQUENCY: { method: 'frequency', sizeOfResultArrayMethod: 'frequencySize', parameters: [RANGE, RANGE] },
    TREND: { method: 'trend', sizeOfResultArrayMethod: 'trendSize', parameters: [RANGE, { argumentType: A.RANGE, optionalArg: true }, { argumentType: A.RANGE, optionalArg: true }] },
    FORECAST: { method: 'forecast', parameters: [NUM, RANGE, RANGE] },
    'FORECAST.LINEAR': { method: 'forecast', parameters: [NUM, RANGE, RANGE] },
    INTERCEPT: { method: 'intercept', parameters: [RANGE, RANGE] },
    PERCENTRANK: { method: 'percentrank', parameters: [RANGE, NUM, optNum(3)] },
    'PERCENTRANK.INC': { method: 'percentrank', parameters: [RANGE, NUM, optNum(3)] },
    NUMBERVALUE: { method: 'numbervalue', parameters: [{ argumentType: A.STRING }, { argumentType: A.STRING, optionalArg: true, defaultValue: '.' }, { argumentType: A.STRING, optionalArg: true, defaultValue: ',' }] },
    TYPE: { method: 'typeOf', parameters: [{ argumentType: A.SCALAR }] }, // SCALAR (like ISERROR) lets error values through
    VALUETOTEXT: { method: 'valuetotext', parameters: [ANY, optNum(0)] },
    TEXTSPLIT: { method: 'textsplit', sizeOfResultArrayMethod: 'textsplitSize', parameters: [{ argumentType: A.STRING }, { argumentType: A.ANY, optionalArg: true }, { argumentType: A.ANY, optionalArg: true }, optNum(0), optNum(0), { argumentType: A.ANY, optionalArg: true }] },
    WRAPROWS: { method: 'wraprows', sizeOfResultArrayMethod: 'wraprowsSize', parameters: [RANGE, NUM, { argumentType: A.ANY, optionalArg: true }] },
    WRAPCOLS: { method: 'wrapcols', sizeOfResultArrayMethod: 'wrapcolsSize', parameters: [RANGE, NUM, { argumentType: A.ANY, optionalArg: true }] },
    EXPAND: { method: 'expand', sizeOfResultArrayMethod: 'expandSize', parameters: [RANGE, NUM, { argumentType: A.NUMBER, optionalArg: true }, { argumentType: A.ANY, optionalArg: true }] },
    RANDARRAY: { method: 'randarray', sizeOfResultArrayMethod: 'randarraySize', isVolatile: true, parameters: [optNum(1), optNum(1), optNum(0), optNum(1), optNum(0)] },
    ARRAYTOTEXT: { method: 'arraytotext', parameters: [RANGE, optNum(0)] },
    CELL: { method: 'cellInfo', parameters: [{ argumentType: A.ANY }, { argumentType: A.ANY, optionalArg: true }] },
    INFO: { method: 'info', parameters: [{ argumentType: A.STRING }] },
    LET: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    LAMBDA: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    MAP: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    REDUCE: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    SCAN: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    BYROW: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    BYCOL: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false },
    MAKEARRAY: { method: 'unexpanded', parameters: [{ argumentType: A.ANY }], repeatLastArgs: 1, expandRanges: false }
  }

  // ---- result sizes (from the formula text) ----
  private sizeOf(ast: Any, i: number): [number, number] | null { return dims(ast.args?.[i]) }
  sameSize(ast: Any): ArraySize { const d = this.sizeOf(ast, 0); return d ? new ArraySize(d[1], d[0]) : bad() }
  takeSize(ast: Any): ArraySize {
    const d = this.sizeOf(ast, 0), r = lit(ast.args[1]), c = ast.args[2] ? lit(ast.args[2]) : null
    if (!d || r === null || (ast.args[2] && c === null)) return bad()
    return new ArraySize(ast.args[2] ? Math.min(Math.abs(c!), d[1]) : d[1], Math.min(Math.abs(r), d[0]))
  }
  dropSize(ast: Any): ArraySize {
    const d = this.sizeOf(ast, 0), r = lit(ast.args[1]), c = ast.args[2] ? lit(ast.args[2]) : null
    if (!d || r === null || (ast.args[2] && c === null)) return bad()
    return new ArraySize(Math.max(1, d[1] - (ast.args[2] ? Math.abs(c!) : 0)), Math.max(1, d[0] - Math.abs(r)))
  }
  choosecolsSize(ast: Any): ArraySize { const d = this.sizeOf(ast, 0); return d ? new ArraySize(ast.args.length - 1, d[0]) : bad() }
  chooserowsSize(ast: Any): ArraySize { const d = this.sizeOf(ast, 0); return d ? new ArraySize(d[1], ast.args.length - 1) : bad() }
  tocolSize(ast: Any): ArraySize { const d = this.sizeOf(ast, 0); const ig = ast.args[1] ? lit(ast.args[1]) : 0; return d && ig === 0 ? new ArraySize(1, d[0] * d[1]) : bad() }
  torowSize(ast: Any): ArraySize { const d = this.sizeOf(ast, 0); const ig = ast.args[1] ? lit(ast.args[1]) : 0; return d && ig === 0 ? new ArraySize(d[0] * d[1], 1) : bad() }
  sortbySize(ast: Any): ArraySize { return this.sameSize(ast) }
  munitSize(ast: Any): ArraySize { const n = lit(ast.args[0]); return n !== null && n >= 1 ? new ArraySize(Math.trunc(n), Math.trunc(n)) : bad() }
  frequencySize(ast: Any): ArraySize { const d = this.sizeOf(ast, 1); return d ? new ArraySize(1, d[0] * d[1] + 1) : bad() }
  trendSize(ast: Any): ArraySize {
    const d = this.sizeOf(ast, ast.args[2] ? 2 : 0)
    return d ? new ArraySize(d[1], d[0]) : bad()
  }

  // ---- array shaping ----
  take(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('TAKE'), (arr: Any, rows: number, cols?: number) => {
      const g = grid(arr), r = Math.trunc(rows)
      let res = r >= 0 ? g.slice(0, r) : g.slice(g.length + r)
      if (cols !== undefined) { const c = Math.trunc(cols); res = res.map((row) => (c >= 0 ? row.slice(0, c) : row.slice(row.length + c))) }
      return res.length && res[0].length ? out(res) : VALUE()
    })
  }
  drop(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('DROP'), (arr: Any, rows: number, cols?: number) => {
      const g = grid(arr), r = Math.trunc(rows)
      let res = r >= 0 ? g.slice(r) : g.slice(0, g.length + r)
      if (cols !== undefined) { const c = Math.trunc(cols); res = res.map((row) => (c >= 0 ? row.slice(c) : row.slice(0, row.length + c))) }
      return res.length && res[0].length ? out(res) : VALUE()
    })
  }
  choosecols(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('CHOOSECOLS'), (arr: Any, ...idx: number[]) => {
      const g = grid(arr), w = g[0]?.length ?? 0
      const cols = idx.map((i) => { const k = Math.trunc(i); return k > 0 ? k - 1 : k < 0 ? w + k : -1 })
      if (cols.some((c) => c < 0 || c >= w)) return VALUE()
      return out(g.map((row) => cols.map((c) => row[c])))
    })
  }
  chooserows(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('CHOOSEROWS'), (arr: Any, ...idx: number[]) => {
      const g = grid(arr), h = g.length
      const rows = idx.map((i) => { const k = Math.trunc(i); return k > 0 ? k - 1 : k < 0 ? h + k : -1 })
      if (rows.some((r) => r < 0 || r >= h)) return VALUE()
      return out(rows.map((r) => g[r]))
    })
  }
  private flat(g: unknown[][], byCol: boolean): unknown[] {
    if (!byCol) return g.flat()
    const res: unknown[] = []
    for (let c = 0; c < (g[0]?.length ?? 0); c++) for (const row of g) res.push(row[c])
    return res
  }
  tocol(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('TOCOL'), (arr: Any, _ignore: number, byCol: boolean) => out(this.flat(grid(arr), byCol).map((v) => [v])))
  }
  torow(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('TOROW'), (arr: Any, _ignore: number, byCol: boolean) => out([this.flat(grid(arr), byCol)]))
  }
  sortby(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('SORTBY'), (arr: Any, ...rest: unknown[]) => {
      const g = grid(arr)
      const keys: { vals: unknown[]; dir: number }[] = []
      for (let i = 0; i < rest.length; i += 2) {
        const vals = grid(rest[i]).flat()
        if (vals.length !== g.length) return VALUE()
        keys.push({ vals, dir: Math.trunc(Number(rest[i + 1] ?? 1)) === -1 ? -1 : 1 })
      }
      const rank = (v: unknown): [number, number | string] => (typeof v === 'number' ? [0, v] : typeof v === 'string' ? [1, v.toLowerCase()] : typeof v === 'boolean' ? [2, Number(v)] : [3, 0])
      const order = g.map((_r, i) => i).sort((a, b) => {
        for (const k of keys) {
          const [ta, va] = rank(k.vals[a]), [tb, vb] = rank(k.vals[b])
          const c = ta !== tb ? ta - tb : va < vb ? -1 : va > vb ? 1 : 0
          if (c) return c * k.dir
        }
        return a - b
      })
      return out(order.map((i) => g[i]))
    })
  }

  // ---- matrices ----
  munit(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('MUNIT'), (n: number) => {
      const k = Math.trunc(n)
      return k >= 1 ? out(Array.from({ length: k }, (_, i) => Array.from({ length: k }, (_, j) => (i === j ? 1 : 0)))) : VALUE()
    })
  }
  minverse(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('MINVERSE'), (arr: Any) => {
      const m = square(grid(arr))
      if (!m) return VALUE()
      const r = luInverse(m)
      return r ? out(r.inv) : new CellError(ErrorType.NUM)
    })
  }
  mdeterm(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('MDETERM'), (arr: Any) => { const m = square(grid(arr)); return m ? determinant(m) : VALUE() })
  }

  // ---- statistics ----
  frequency(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('FREQUENCY'), (data: Any, bins: Any) => {
      const d = grid(data).flat().map(num).filter((x): x is number => x !== null)
      const b = grid(bins).flat().map(num).filter((x): x is number => x !== null).sort((x, y) => x - y)
      const counts = new Array(b.length + 1).fill(0)
      for (const v of d) { let k = b.findIndex((x) => v <= x); if (k < 0) k = b.length; counts[k]++ }
      return out(counts.map((c) => [c]))
    })
  }
  trend(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('TREND'), (ys: Any, xs?: Any, nx?: Any) => {
      const yg = grid(ys), y = yg.flat()
      const x = xs ? grid(xs).flat() : y.map((_v, i) => i + 1)
      const p = pairs([y], [x])
      if (p instanceof CellError) return p
      const f = line(p)
      if (!f) return DIV0()
      const target = nx ? grid(nx) : xs ? grid(xs) : yg.map((row, i) => row.map((_v, j) => i * row.length + j + 1))
      return out(target.map((row) => row.map((v) => { const a = num(v); return a === null ? VALUE() : f.intercept + f.slope * a })))
    })
  }
  forecast(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('FORECAST'), (x: number, ys: Any, xs: Any) => {
      const p = pairs(grid(ys), grid(xs))
      if (p instanceof CellError) return p
      const f = line(p)
      return f ? f.intercept + f.slope * x : DIV0()
    })
  }
  intercept(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('INTERCEPT'), (ys: Any, xs: Any) => {
      const p = pairs(grid(ys), grid(xs))
      if (p instanceof CellError) return p
      const f = line(p)
      return f ? f.intercept : DIV0()
    })
  }
  percentrank(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('PERCENTRANK'), (arr: Any, x: number, sig: number) => {
      const v = grid(arr).flat().map(num).filter((n): n is number => n !== null).sort((a, b) => a - b)
      const n = v.length
      if (!n || x < v[0] || x > v[n - 1]) return NA()
      let rank: number
      const exact = v.indexOf(x)
      if (exact >= 0) rank = exact / (n - 1)
      else {
        let i = 0
        while (v[i + 1] < x) i++
        rank = (i + (x - v[i]) / (v[i + 1] - v[i])) / (n - 1)
      }
      const d = Math.max(1, Math.trunc(sig)), f = 10 ** d
      return Math.floor(rank * f) / f
    })
  }

  // ---- scalars ----
  numbervalue(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('NUMBERVALUE'), (text: string, dec: string, grp: string) => {
      let t = String(text).replace(/\s+/g, '')
      if (t === '') return 0
      let pct = 0
      while (t.endsWith('%')) { pct++; t = t.slice(0, -1) }
      if (grp) t = t.split(grp[0]).join('')
      if (dec && dec[0] !== '.') t = t.split(dec[0]).join('.')
      if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return VALUE()
      return Number(t) / 100 ** pct
    })
  }
  typeOf(ast: Any, state: Any): Any {
    const a = ast.args?.[0]
    if (a?.start && a?.end && (a.start.row !== a.end.row || a.start.col !== a.end.col)) return 64 // a multi-cell range is an array
    return this.runFunction(ast.args, state, this.metadata('TYPE'), (v: unknown) => {
      if (v instanceof SimpleRangeValue) return 64
      if (v instanceof CellError) return 16
      if (typeof v === 'boolean') return 4
      if (typeof v === 'string') return 2
      return 1
    })
  }
  valuetotext(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('VALUETOTEXT'), (v: unknown, fmt: number) => {
      if (v instanceof CellError) return '#VALUE!'
      const n = num(v)
      if (typeof v === 'string') return Math.trunc(fmt) === 1 ? '"' + v.replace(/"/g, '""') + '"' : v
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
      return n === null ? '' : String(n)
    })
  }
  // ---- text splitting and array reshaping ----
  /** A string the formula text pins down: a literal, or the current value of a referenced cell (the size is fixed when the formula is entered). */
  private peek(node: Any, state: Any): unknown {
    if (!node) return undefined
    if (node.type === 'STRING') return node.value
    if (node.type === 'CELL_REFERENCE') { try { return this.dependencyGraph.getScalarValue(node.reference.toSimpleCellAddress(state.formulaAddress)) } catch { return undefined } }
    return undefined
  }
  private delims(node: Any, state: Any): string[] | null {
    if (!node || node.type === 'EMPTY') return []
    const v = this.peek(node, state)
    return typeof v === 'string' ? (v === '' ? [] : [v]) : null
  }
  private split(text: string, cols: string[], rows: string[], ignoreEmpty: boolean, insensitive: boolean): string[][] {
    const re = (d: string[]): RegExp | null => (d.length ? new RegExp(d.map((x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), insensitive ? 'gi' : 'g') : null)
    const rr = re(rows), cr = re(cols)
    let lines = rr ? text.split(rr) : [text]
    if (ignoreEmpty) lines = lines.filter((l) => l !== '')
    const g = lines.map((l) => { const parts = cr ? l.split(cr) : [l]; return ignoreEmpty ? parts.filter((x) => x !== '') : parts })
    return g.length ? g : [['']]
  }
  textsplitSize(ast: Any, state: Any): ArraySize {
    const t = this.peek(ast.args[0], state), cd = this.delims(ast.args[1], state), rd = this.delims(ast.args[2], state)
    if (typeof t !== 'string' || !cd || !rd) return bad()
    const ig = ast.args[3] ? lit(ast.args[3]) : 0, mm = ast.args[4] ? lit(ast.args[4]) : 0
    if (ig === null || mm === null) return bad()
    const g = this.split(t, cd, rd, ig !== 0, mm === 1)
    return new ArraySize(Math.max(1, ...g.map((r) => r.length)), g.length)
  }
  textsplit(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('TEXTSPLIT'), (text: string, cd: unknown, rd: unknown, ignore: number, mode: number, pad: unknown) => {
      const flat = (d: unknown): string[] => (d === undefined || d === null ? [] : grid(d).flat().map(String).filter((x) => x !== ''))
      const g = this.split(text, flat(cd), flat(rd), ignore !== 0, mode === 1)
      const w = Math.max(1, ...g.map((r) => r.length))
      const fill = pad === undefined || pad === null ? NA() : pad
      return out(g.map((r) => Array.from({ length: w }, (_v, k) => (k < r.length ? r[k] : fill))))
    })
  }
  private wrapDims(ast: Any, byCol: boolean): ArraySize {
    const d = this.sizeOf(ast, 0), n = lit(ast.args[1])
    if (!d || n === null || n < 1) return bad()
    const len = d[0] * d[1], k = Math.trunc(n), other = Math.ceil(len / k)
    return byCol ? new ArraySize(other, k) : new ArraySize(k, other)
  }
  wraprowsSize(ast: Any): ArraySize { return this.wrapDims(ast, false) }
  wrapcolsSize(ast: Any): ArraySize { return this.wrapDims(ast, true) }
  private wrap(ast: Any, state: Any, name: string, byCol: boolean): Any {
    return this.runFunction(ast.args, state, this.metadata(name), (arr: Any, n: number, pad: unknown) => {
      const v = grid(arr).flat(), k = Math.trunc(n)
      if (k < 1) return VALUE()
      const other = Math.ceil(v.length / k)
      const fill = pad === undefined || pad === null ? NA() : pad
      const at = (i: number): unknown => (i < v.length ? v[i] : fill)
      return out(byCol
        ? Array.from({ length: k }, (_a, r) => Array.from({ length: other }, (_b, c) => at(c * k + r)))
        : Array.from({ length: other }, (_a, r) => Array.from({ length: k }, (_b, c) => at(r * k + c))))
    })
  }
  wraprows(ast: Any, state: Any): Any { return this.wrap(ast, state, 'WRAPROWS', false) }
  wrapcols(ast: Any, state: Any): Any { return this.wrap(ast, state, 'WRAPCOLS', true) }
  expandSize(ast: Any): ArraySize {
    const d = this.sizeOf(ast, 0), r = lit(ast.args[1]), c = ast.args[2] ? lit(ast.args[2]) : d ? d[1] : null
    return d && r !== null && c !== null && r >= d[0] && c >= d[1] ? new ArraySize(Math.trunc(c), Math.trunc(r)) : bad()
  }
  expand(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('EXPAND'), (arr: Any, rows: number, cols: number | undefined, pad: unknown) => {
      const g = grid(arr), r = Math.trunc(rows), c = cols === undefined ? g[0].length : Math.trunc(cols)
      if (r < g.length || c < g[0].length) return VALUE()
      const fill = pad === undefined || pad === null ? NA() : pad
      return out(Array.from({ length: r }, (_a, i) => Array.from({ length: c }, (_b, j) => (i < g.length && j < g[0].length ? g[i][j] : fill))))
    })
  }
  randarraySize(ast: Any): ArraySize {
    const r = ast.args[0] ? lit(ast.args[0]) : 1, c = ast.args[1] ? lit(ast.args[1]) : 1
    return r !== null && c !== null && r >= 1 && c >= 1 ? new ArraySize(Math.trunc(c), Math.trunc(r)) : bad()
  }
  randarray(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('RANDARRAY'), (rows: number, cols: number, min: number, max: number, whole: number) => {
      if (min > max) return VALUE()
      const lo = Math.ceil(min), hi = Math.floor(max)
      if (whole && lo > hi) return VALUE()
      return out(Array.from({ length: Math.trunc(rows) }, () => Array.from({ length: Math.trunc(cols) }, () => (whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : min + Math.random() * (max - min)))))
    })
  }
  arraytotext(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('ARRAYTOTEXT'), (arr: Any, fmt: number) => {
      const strict = Math.trunc(fmt) === 1
      const show = (v: unknown): string => (v instanceof CellError ? '#VALUE!' : typeof v === 'string' ? (strict ? '"' + v.replace(/"/g, '""') + '"' : v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v === null || v === undefined ? '' : String(num(v) ?? v))
      const g = grid(arr)
      return strict ? '{' + g.map((r) => r.map(show).join(',')).join(';') + '}' : g.map((r) => r.map(show).join(', ')).join(', ')
    })
  }

  /** LET / LAMBDA-family calls are expanded before they reach the engine (engine/lambda.ts); one that arrives here couldn't be (size not literal). */
  unexpanded(): Any { return new CellError(ErrorType.NAME) }

  // ---- CELL and INFO ----
  cellInfo(ast: Any, state: Any): Any {
    const kind = this.peek(ast.args[0], state)
    if (typeof kind !== 'string') return VALUE()
    const node = ast.args[1]
    const ref = node ? (node.reference ?? node.start) : null
    if (node && !ref) return VALUE()
    const a = ref ? ref.toSimpleCellAddress(state.formulaAddress) : state.formulaAddress
    const colLetters = (c: number): string => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s }
    const v = this.dependencyGraph.getScalarValue(a) as unknown
    switch (kind.toLowerCase()) {
      case 'address': return `$${colLetters(a.col)}$${a.row + 1}`
      case 'col': return a.col + 1
      case 'row': return a.row + 1
      case 'contents': return v === null || v === undefined ? 0 : (v as Any)
      case 'type': return v === null || v === undefined || v === '' ? 'b' : typeof v === 'string' ? 'l' : 'v'
      case 'filename': case 'prefix': return ''
      case 'protect': case 'parentheses': case 'color': return 0
      default: return VALUE()
    }
  }
  info(ast: Any, state: Any): Any {
    return this.runFunction(ast.args, state, this.metadata('INFO'), (what: string) => {
      switch (String(what).toLowerCase()) {
        case 'system': return 'pcdos'
        case 'recalc': return 'Automatic'
        case 'numfile': return 1
        case 'osversion': return 'Windows (32-bit) NT 10.00'
        case 'release': return '16.0'
        case 'origin': return '$A:$A$1'
        case 'directory': return ''
        case 'memavail': case 'memused': case 'totmem': return 0
        default: return VALUE()
      }
    })
  }

}

let registered = false
export function registerMoreFunctions(): void {
  if (registered) return
  registered = true
  const names = Object.fromEntries(Object.keys(MorePlugin.implementedFunctions).map((k) => [k, k]))
  HyperFormula.registerFunctionPlugin(MorePlugin, { enGB: names, enUS: names } as never)
}

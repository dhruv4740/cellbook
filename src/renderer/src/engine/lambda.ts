// LAMBDA, MAP, BYROW, BYCOL, REDUCE, SCAN and MAKEARRAY. HyperFormula can't pass a function to a function, so these are expanded
// textually before the formula reaches the engine (like LET): each call is unrolled over its array, which must therefore be a
// literal-sized range, SEQUENCE(...) or {array constant}. Results are assembled with VSTACK / HSTACK. The typed text is kept for display and saving.
import { closeParen, splitArgs, substituteAll, expandLet } from './let'
import { colName, colIndex } from '../model/types'

const MAX_LEN = 20000
const IDENT = /^[A-Za-z_][A-Za-z0-9_.]*$/
const REF = /^((?:'[^']+'|[A-Za-z0-9_.]+)!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/

interface Arr { r: number; c: number; at: (i: number, j: number) => string; row: (i: number) => string; col: (j: number) => string }

const wrap = (v: string): string => (/^-?[\d.]+$/.test(v) || REF.test(v) || /^\{[^{}]*\}$/.test(v) ? v : `(${v})`)

function stripParens(e: string): string {
  e = e.trim()
  while (e[0] === '(' && closeParen(e, 0) === e.length - 1) e = e.slice(1, -1).trim()
  return e
}

function literal(e: string): number | null {
  const t = stripParens(e)
  return /^-?\d+(\.\d+)?$/.test(t) ? +t : null
}

function fromMatrix(m: string[][]): Arr {
  const c = m[0]?.length ?? 0
  return { r: m.length, c, at: (i, j) => m[i][j], row: (i) => `{${m[i].join(',')}}`, col: (j) => `{${m.map((r) => r[j]).join(';')}}` }
}

/** Dimensions and element accessors of a literal-sized array expression; null when the size can't be read from the text. */
function arrayOf(expr: string): Arr | null {
  const e = stripParens(expr)
  const ref = REF.exec(e)
  if (ref) {
    const pre = ref[1] ?? ''
    const c1 = colIndex(ref[2].toUpperCase()), r1 = +ref[3] - 1
    const c2 = ref[4] ? colIndex(ref[4].toUpperCase()) : c1, r2 = ref[5] ? +ref[5] - 1 : r1
    const cell = (r: number, c: number): string => colName(c) + (r + 1)
    return {
      r: r2 - r1 + 1, c: c2 - c1 + 1,
      at: (i, j) => pre + cell(r1 + i, c1 + j),
      row: (i) => (c1 === c2 ? pre + cell(r1 + i, c1) : `${pre}${cell(r1 + i, c1)}:${cell(r1 + i, c2)}`),
      col: (j) => (r1 === r2 ? pre + cell(r1, c1 + j) : `${pre}${cell(r1, c1 + j)}:${cell(r2, c1 + j)}`)
    }
  }
  const seq = /^SEQUENCE\s*\(([\s\S]*)\)$/i.exec(e)
  if (seq && closeParen(e, e.indexOf('(')) === e.length - 1) {
    const raw = splitArgs(seq[1]).map((x) => x.trim())
    const a = raw.map((x) => (x === '' ? undefined : literal(x)))
    if (raw.length > 4 || a.some((x) => x === null) || a[0] === undefined) return null
    const rows = a[0] as number, cols = a[1] ?? 1, start = a[2] ?? 1, step = a[3] ?? 1
    if (rows < 1 || cols < 1 || rows * cols > 5000) return null
    return fromMatrix(Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => String(start + step * (i * cols + j)))))
  }
  if (e[0] === '{' && e[e.length - 1] === '}') {
    const rows = e.slice(1, -1).split(';').map((r) => splitArgs(r).map((x) => x.trim()))
    if (!rows.length || rows.some((r) => r.length !== rows[0].length)) return null
    return fromMatrix(rows)
  }
  return null
}

const stack = (g: string[][]): string => {
  const rows = g.map((r) => (r.length === 1 ? r[0] : `HSTACK(${r.join(',')})`))
  return rows.length === 1 ? rows[0] : `VSTACK(${rows.join(',')})`
}

interface Lam { params: string[]; body: string }
function lambdaOf(text: string): Lam | null {
  const t = text.trim()
  const m = /^LAMBDA\s*\(/i.exec(t)
  if (!m) return null
  const open = m[0].length - 1
  if (closeParen(t, open) !== t.length - 1) return null
  const a = splitArgs(t.slice(open + 1, -1)).map((x) => x.trim())
  if (a.length < 1) return null
  const params = a.slice(0, -1)
  if (params.some((p) => !IDENT.test(p))) return null
  return { params, body: a[a.length - 1] }
}

const apply = (l: Lam, values: string[]): string => substituteAll(l.body, new Map(l.params.map((p, k) => [p.toLowerCase(), wrap(values[k])])))

/** Unrolls one MAP / BYROW / ... call; returns the replacement text or null. */
function unroll(name: string, args: string[]): string | null {
  const lam = lambdaOf(args[args.length - 1] ?? '')
  if (!lam) return null
  const out = (s: string): string | null => (s.length > MAX_LEN ? null : s)
  switch (name) {
    case 'MAP': {
      const arrs = args.slice(0, -1).map(arrayOf)
      if (!arrs.length || arrs.some((a) => !a) || lam.params.length !== arrs.length) return null
      const first = arrs[0]!
      if (arrs.some((a) => a!.r !== first.r || a!.c !== first.c)) return null
      return out(stack(Array.from({ length: first.r }, (_, i) => Array.from({ length: first.c }, (_, j) => apply(lam, arrs.map((a) => a!.at(i, j)))))))
    }
    case 'BYROW': {
      const a = args.length === 2 ? arrayOf(args[0]) : null
      if (!a || lam.params.length !== 1) return null
      return out(stack(Array.from({ length: a.r }, (_, i) => [apply(lam, [a.row(i)])])))
    }
    case 'BYCOL': {
      const a = args.length === 2 ? arrayOf(args[0]) : null
      if (!a || lam.params.length !== 1) return null
      return out(stack([Array.from({ length: a.c }, (_, j) => apply(lam, [a.col(j)]))]))
    }
    case 'MAKEARRAY': {
      const r = args.length === 3 ? literal(args[0]) : null, c = args.length === 3 ? literal(args[1]) : null
      if (r === null || c === null || r < 1 || c < 1 || r * c > 5000 || lam.params.length !== 2) return null
      return out(stack(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => apply(lam, [String(i + 1), String(j + 1)])))))
    }
    case 'REDUCE': case 'SCAN': {
      if (args.length < 2 || args.length > 3 || lam.params.length !== 2) return null
      const a = arrayOf(args[args.length - 2])
      if (!a) return null
      const elems: string[] = []
      for (let i = 0; i < a.r; i++) for (let j = 0; j < a.c; j++) elems.push(a.at(i, j))
      let acc: string
      let k = 0
      if (args.length === 3) acc = args[0]
      else { acc = elems[0]; k = 1 }
      const trail: string[] = []
      for (; k < elems.length; k++) {
        acc = apply(lam, [acc, elems[k]])
        if (acc.length > MAX_LEN) return null
        trail.push(acc)
      }
      if (name === 'REDUCE') return acc
      if (args.length === 2) trail.unshift(elems[0])
      return out(stack(Array.from({ length: a.r }, (_, i) => Array.from({ length: a.c }, (_, j) => trail[i * a.c + j]))))
    }
  }
  return null
}

const CALL = /(?<![A-Za-z0-9_.])(MAP|BYROW|BYCOL|REDUCE|SCAN|MAKEARRAY|LAMBDA)\s*\(/gi

/** Expands every LAMBDA-family call in a formula; null when it has none or one can't be unrolled. */
export function expandLambdas(formula: string): string | null {
  if (!/(?<![A-Za-z0-9_.])(MAP|BYROW|BYCOL|REDUCE|SCAN|MAKEARRAY|LAMBDA)\s*\(/i.test(formula)) return null
  let f = formula
  for (let guard = 0; guard < 400; guard++) {
    CALL.lastIndex = 0
    const m = CALL.exec(f)
    if (!m) return f
    const name = m[1].toUpperCase()
    const open = m.index + m[0].length - 1
    const close = closeParen(f, open)
    if (close < 0) return null
    let rep: string | null
    let end = close + 1
    if (name === 'LAMBDA') {
      const lam = lambdaOf(f.slice(m.index, close + 1))
      let p = close + 1
      while (f[p] === ' ') p++
      if (!lam || f[p] !== '(') return null
      const callClose = closeParen(f, p)
      if (callClose < 0) return null
      const callArgs = splitArgs(f.slice(p + 1, callClose)).map((x) => x.trim())
      if (callArgs.length !== lam.params.length) return null
      rep = apply(lam, callArgs)
      end = callClose + 1
    } else {
      rep = unroll(name, splitArgs(f.slice(open + 1, close)).map((x) => x.trim()))
    }
    if (rep === null || f.length - (end - m.index) + rep.length > MAX_LEN * 2) return null
    f = f.slice(0, m.index) + (/^(VSTACK|HSTACK)\(/.test(rep) ? rep : `(${rep})`) + f.slice(end)
  }
  return null
}

/** LET and LAMBDA expansion: the formula text the engine should hold, or null when it needs none. */
export function expandFormula(formula: string): string | null {
  const a = expandLet(formula)
  const b = expandLambdas(a ?? formula)
  return b ?? a
}

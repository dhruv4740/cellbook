// Flash Fill (Ctrl+E): learns a string transformation from the examples typed next to the data and applies it to the
// remaining rows. A small program synthesiser in the spirit of Gulwani's FlashFill: an output is a concatenation of
// constants and substrings of the input cells, with positions found by delimiters / word boundaries / offsets.

type Case = 'none' | 'lower' | 'upper' | 'proper'
type TokenClass = 'word' | 'alpha' | 'digit'
type Pos =
  | { k: 'abs'; n: number }
  | { k: 'end'; n: number }
  | { k: 'delim'; ch: string; n: number; side: 'before' | 'after' }
  | { k: 'tok'; cls: TokenClass; n: number; side: 'start' | 'end' }
type Atom = { t: 'const'; s: string } | { t: 'sub'; col: number; l: Pos; r: Pos; cs: Case }

const isAlpha = (ch: string): boolean => /\p{L}/u.test(ch)
const isDigit = (ch: string): boolean => /[0-9]/.test(ch)
const isWordCh = (ch: string): boolean => isAlpha(ch) || isDigit(ch)
const CLASS: Record<TokenClass, (ch: string) => boolean> = { word: isWordCh, alpha: isAlpha, digit: isDigit }

function tokens(s: string, cls: TokenClass): [number, number][] {
  const out: [number, number][] = []
  const f = CLASS[cls]
  for (let i = 0; i < s.length; ) {
    if (!f(s[i])) { i++; continue }
    let j = i
    while (j < s.length && f(s[j])) j++
    out.push([i, j])
    i = j
  }
  return out
}
function delimIdx(s: string, ch: string): number[] {
  const out: number[] = []
  for (let i = 0; i < s.length; i++) if (s[i] === ch) out.push(i)
  return out
}
function pick<T>(list: T[], n: number): T | undefined { return n > 0 ? list[n - 1] : list[list.length + n] }

function evalPos(s: string, p: Pos): number | null {
  switch (p.k) {
    case 'abs': return p.n <= s.length ? p.n : null
    case 'end': return s.length - p.n >= 0 ? s.length - p.n : null
    case 'delim': { const i = pick(delimIdx(s, p.ch), p.n); return i === undefined ? null : p.side === 'before' ? i : i + 1 }
    case 'tok': { const t = pick(tokens(s, p.cls), p.n); return t === undefined ? null : p.side === 'start' ? t[0] : t[1] }
  }
}

function posSpecs(s: string, at: number): Pos[] {
  const out: Pos[] = [{ k: 'abs', n: at }, { k: 'end', n: s.length - at }]
  for (const cls of ['word', 'alpha', 'digit'] as TokenClass[]) {
    const ts = tokens(s, cls)
    ts.forEach((t, i) => {
      if (t[0] === at) { out.push({ k: 'tok', cls, n: i + 1, side: 'start' }); out.push({ k: 'tok', cls, n: i - ts.length, side: 'start' }) }
      if (t[1] === at) { out.push({ k: 'tok', cls, n: i + 1, side: 'end' }); out.push({ k: 'tok', cls, n: i - ts.length, side: 'end' }) }
    })
  }
  const seen = new Set<string>()
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (isWordCh(ch) || seen.has(ch)) continue
    seen.add(ch)
    const idx = delimIdx(s, ch)
    idx.forEach((x, i2) => {
      const side = x === at ? 'before' : x + 1 === at ? 'after' : null
      if (side) { out.push({ k: 'delim', ch, n: i2 + 1, side }); out.push({ k: 'delim', ch, n: i2 - idx.length, side }) }
    })
  }
  return out
}

const posCost = (p: Pos): number => (p.k === 'abs' ? (p.n === 0 ? 0.2 : 2.5) : p.k === 'end' ? (p.n === 0 ? 0.2 : 2.5) : 0.5 + (Math.abs(p.n) > 1 ? 0.1 : 0))

function applyCase(s: string, cs: Case): string {
  switch (cs) {
    case 'lower': return s.toLowerCase()
    case 'upper': return s.toUpperCase()
    case 'proper': return s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase())
    default: return s
  }
}

function evalAtom(a: Atom, row: string[]): string | null {
  if (a.t === 'const') return a.s
  const s = row[a.col] ?? ''
  const l = evalPos(s, a.l), r = evalPos(s, a.r)
  if (l === null || r === null || l >= r) return null
  return applyCase(s.slice(l, r), a.cs)
}
let inputChars = new Set<string>()
const atomCost = (a: Atom): number => (a.t === 'const' ? 3 + 0.5 * a.s.length + 6 * [...a.s].filter((ch) => isWordCh(ch) && inputChars.has(ch.toLowerCase())).length : 1 + posCost(a.l) + posCost(a.r) + (a.cs === 'none' ? 0 : 0.5))

interface Example { row: string[]; out: string }

/** Substring atoms of the first example that could produce some part of its output. */
function candidateAtoms(ex: Example): Atom[] {
  const atoms: Atom[] = []
  const out = ex.out
  ex.row.forEach((s, col) => {
    if (!s) return
    for (let p = 0; p < s.length; p++) {
      for (let q = p + 1; q <= s.length; q++) {
        const sub = s.slice(p, q)
        for (const cs of ['none', 'lower', 'upper', 'proper'] as Case[]) {
          const v = applyCase(sub, cs)
          if (cs !== 'none' && v === sub) continue
          if (!out.includes(v)) continue
          const ls = posSpecs(s, p), rs = posSpecs(s, q)
          for (const l of ls) for (const r of rs) atoms.push({ t: 'sub', col, l, r, cs })
        }
      }
    }
  })
  return atoms
}

function synthesise(examples: Example[]): Atom[] | null {
  const first = examples[0]
  inputChars = new Set(first.row.join('').toLowerCase())
  const subs = candidateAtoms(first)
  // atom -> evaluation per example (null when it does not apply)
  const evals = subs.map((a) => examples.map((e) => evalAtom(a, e.row)))
  const byStart = new Map<number, number[]>() // position in the first output -> atom indexes
  subs.forEach((_a, i) => {
    const v = evals[i][0]
    if (!v || evals[i].some((x) => x === null || x === '')) return
    for (let p = first.out.indexOf(v); p >= 0; p = first.out.indexOf(v, p + 1)) {
      let a = byStart.get(p)
      if (!a) byStart.set(p, (a = []))
      a.push(i)
    }
  })
  const n = examples.length
  const key = (st: number[]): string => st.join(',')
  const goal = examples.map((e) => e.out.length)
  const start = examples.map(() => 0)
  const dist = new Map<string, number>([[key(start), 0]])
  const prev = new Map<string, { from: number[]; atom: Atom }>()
  const heap: { st: number[]; d: number }[] = [{ st: start, d: 0 }]
  const done = new Set<string>()
  let expansions = 0
  while (heap.length) {
    let bi = 0
    for (let i = 1; i < heap.length; i++) if (heap[i].d < heap[bi].d) bi = i
    const { st, d } = heap.splice(bi, 1)[0]
    const k = key(st)
    if (done.has(k)) continue
    done.add(k)
    if (k === key(goal)) {
      const path: Atom[] = []
      for (let cur = k; cur !== key(start); ) { const p = prev.get(cur)!; path.unshift(p.atom); cur = key(p.from) }
      return path
    }
    if (++expansions > 4000) return null
    const relax = (atom: Atom, next: number[]): void => {
      const nk = key(next)
      const nd = d + atomCost(atom)
      if (nd < (dist.get(nk) ?? Infinity)) { dist.set(nk, nd); prev.set(nk, { from: st, atom }); heap.push({ st: next, d: nd }) }
    }
    // Sub atoms
    for (const i of byStart.get(st[0]) ?? []) {
      const next: number[] = []
      let ok = true
      for (let e = 0; e < n; e++) {
        const v = evals[i][e]!
        if (!examples[e].out.startsWith(v, st[e])) { ok = false; break }
        next.push(st[e] + v.length)
      }
      if (ok) relax(subs[i], next)
    }
    // Constants: every prefix of the remaining first output that all outputs share at their positions
    const rest = first.out.slice(st[0])
    for (let len = 1; len <= rest.length; len++) {
      const c = rest.slice(0, len)
      if (examples.every((e, i) => e.out.startsWith(c, st[i]))) relax({ t: 'const', s: c }, st.map((x) => x + len))
    }
  }
  return null
}

function run(program: Atom[], row: string[]): string | null {
  let out = ''
  for (const a of program) {
    const v = evalAtom(a, row)
    if (v === null) return null
    out += v
  }
  return out
}

/**
 * `rows[i]` are the source cells of row i, `targets[i]` what the user typed in the target column ('' = empty).
 * Returns the values to write for the empty target rows (null = leave empty), or null when no pattern fits.
 */
export function flashFill(rows: string[][], targets: string[]): (string | null)[] | null {
  const idx = targets.map((t, i) => (t !== '' ? i : -1)).filter((i) => i >= 0)
  const attempt = (use: number[]): Atom[] | null => (use.length ? synthesise(use.slice(0, 8).map((i) => ({ row: rows[i], out: targets[i] }))) : null)
  let program = attempt(idx)
  // The first row is usually a header (its "example" is the column title): retry without it.
  if (!program && idx.length > 1) program = attempt(idx.slice(1))
  if (!program) return null
  return targets.map((t, i) => (t !== '' ? null : rows[i].every((c) => c === '') ? null : run(program!, rows[i])))
}

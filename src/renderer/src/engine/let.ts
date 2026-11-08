// LET(name1, value1, ..., body): HyperFormula has no scoped names, so LET formulas are expanded before they
// reach the engine, substituting each name with its (parenthesised) value. The original text is kept for display and saving.

/** Index of the parenthesis closing the one at `open`, skipping string literals; -1 if unbalanced. */
export function closeParen(s: string, open: number): number {
  let depth = 0
  for (let i = open; i < s.length; i++) {
    const ch = s[i]
    if (ch === '"') { i++; while (i < s.length && !(s[i] === '"' && s[i + 1] !== '"')) i += s[i] === '"' ? 2 : 1; continue }
    if (ch === '(') depth++
    else if (ch === ')' && --depth === 0) return i
  }
  return -1
}

/** Splits at top-level commas (not inside parentheses, braces or strings). */
export function splitArgs(s: string): string[] {
  const out: string[] = []
  let depth = 0, start = 0
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (ch === '"') { i++; while (i < s.length && !(s[i] === '"' && s[i + 1] !== '"')) i += s[i] === '"' ? 2 : 1; continue }
    if (ch === '(' || ch === '{') depth++
    else if (ch === ')' || ch === '}') depth--
    else if (ch === ',' && depth === 0) { out.push(s.slice(start, i)); start = i + 1 }
  }
  out.push(s.slice(start))
  return out
}

/** Replaces whole-identifier occurrences of the names in `map` (keys lower-case; outside strings and sheet/cell qualifiers) in one pass. */
export function substituteAll(s: string, map: Map<string, string>): string {
  let out = ''
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (ch === '"') {
      let j = i + 1
      while (j < s.length && !(s[j] === '"' && s[j + 1] !== '"')) j += s[j] === '"' ? 2 : 1
      out += s.slice(i, j + 1); i = j; continue
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i
      while (j < s.length && /[A-Za-z0-9_.]/.test(s[j])) j++
      const word = s.slice(i, j)
      const prev = out[out.length - 1]
      const val = map.get(word.toLowerCase())
      out += val !== undefined && !(prev && /[!$'.]/.test(prev)) && s[j] !== '(' && s[j] !== '!' ? val : word
      i = j - 1; continue
    }
    out += ch
  }
  return out
}
const substitute = (s: string, name: string, value: string): string => substituteAll(s, new Map([[name.toLowerCase(), value]]))

const IDENT = /^[A-Za-z_][A-Za-z0-9_.]*$/

/** Expands every LET in a formula; returns null when the formula has none (or it is malformed). */
export function expandLet(formula: string): string | null {
  if (!/\bLET\s*\(/i.test(formula)) return null
  let f = formula
  for (let guard = 0; guard < 50; guard++) {
    // Innermost LET first: the last "LET(" in the text has no LET inside its arguments.
    const matches = [...f.matchAll(/(?<![A-Za-z0-9_.])LET\s*\(/gi)]
    if (!matches.length) return f
    const m = matches[matches.length - 1]
    const open = m.index! + m[0].length - 1
    const close = closeParen(f, open)
    if (close < 0) return null
    const args = splitArgs(f.slice(open + 1, close)).map((a) => a.trim())
    if (args.length < 3 || args.length % 2 === 0) return null
    const pairs: [string, string][] = []
    for (let i = 0; i < args.length - 1; i += 2) {
      if (!IDENT.test(args[i])) return null
      pairs.push([args[i], args[i + 1]])
    }
    let body = args[args.length - 1]
    for (let k = pairs.length - 1; k >= 0; k--) body = substitute(body, pairs[k][0], `(${pairs[k][1]})`)
    f = f.slice(0, m.index!) + `(${body})` + f.slice(close + 1)
  }
  return null
}

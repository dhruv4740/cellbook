// Home > Increase / Decrease Decimal: edits the digits after the decimal point of every section of
// a number format, keeping currency symbols, colours, text and grouping (like Excel).

function splitSections(fmt: string): string[] {
  const out: string[] = []
  let cur = '', q = false, br = false
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]
    if (ch === '\\' && !q) { cur += ch + (fmt[i + 1] ?? ''); i++; continue }
    if (ch === '"') q = !q
    else if (!q && ch === '[') br = true
    else if (!q && ch === ']') br = false
    if (ch === ';' && !q && !br) { out.push(cur); cur = '' } else cur += ch
  }
  out.push(cur)
  return out
}

/** Index ranges of `fmt` that are literal (quoted, escaped, or [..]) and must not be touched. */
function literalMask(s: string): boolean[] {
  const m = new Array(s.length).fill(false)
  let q = false, br = false
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (ch === '\\' && !q) { m[i] = true; if (i + 1 < s.length) m[i + 1] = true; i++; continue }
    if (ch === '"') { q = !q; m[i] = true; continue }
    if (!q && ch === '[') br = true
    m[i] = q || br
    if (!q && ch === ']') br = false
  }
  return m
}

function adjustSection(sec: string, delta: number): string {
  const lit = literalMask(sec)
  // Date/time or text-only sections are left alone.
  const bare = [...sec].filter((_, i) => !lit[i]).join('')
  if (/[dmyhs]/i.test(bare.replace(/E[+-]/gi, '')) || !/[0#?]/.test(bare)) return sec
  let dot = -1
  for (let i = 0; i < sec.length; i++) if (!lit[i] && sec[i] === '.') { dot = i; break }
  if (dot >= 0) {
    let end = dot + 1
    while (end < sec.length && !lit[end] && /[0#?]/.test(sec[end])) end++
    const decs = end - dot - 1
    if (delta > 0) return sec.slice(0, end) + '0' + sec.slice(end)
    if (decs <= 1) return sec.slice(0, dot) + sec.slice(end) // drop the point with the last digit
    return sec.slice(0, end - 1) + sec.slice(end)
  }
  if (delta < 0) return sec
  let last = -1
  for (let i = 0; i < sec.length; i++) if (!lit[i] && /[0#?]/.test(sec[i])) last = i
  // Stop before an exponent: 0E+00 -> 0.0E+00
  const e = sec.slice(0, last + 1).search(/[Ee][+-]/)
  if (e >= 0) { let k = e - 1; while (k >= 0 && lit[k]) k--; last = k }
  return sec.slice(0, last + 1) + '.0' + sec.slice(last + 1)
}

/**
 * The new format after one Increase (+1) or Decrease (-1) Decimal click. For General, Excel
 * starts from the decimals currently displayed (e.g. 3.14159 shown as 3.14159 -> 0.000000).
 */
export function adjustDecimals(fmt: string | undefined, displayText: string, delta: 1 | -1): string {
  if (!fmt || fmt === 'General') {
    const m = /[.,](\d+)(?:E|%|$)/i.exec(displayText.replace(/\s/g, ''))
    const shown = /\./.test(displayText) && m ? m[1].length : 0
    const pct = /%$/.test(displayText.trim())
    const dec = Math.max(0, shown + delta)
    return '0' + (dec > 0 ? '.' + '0'.repeat(dec) : '') + (pct ? '%' : '')
  }
  return splitSections(fmt).map((s) => adjustSection(s, delta)).join(';')
}

import { SHORT_DATE_FORMAT } from '../engine/locale'
// Excel number-format engine (en-US). Dependency-free implementation of the Format-Cells
// number-format grammar: sections, conditions, colors, digit placeholders, percent, scientific,
// fractions, dates/times and text placeholders. Parsed format codes are cached.
import type { CellValue } from '../model/types'
import { isError } from '../model/types'

export interface FormattedValue {
  text: string
  color?: string // from [Red] etc. in the format code
  align?: 'left' | 'right' | 'center' // natural alignment when hAlign is 'general'
}

// ---------------------------------------------------------------------------------------------
// Colors
// ---------------------------------------------------------------------------------------------

const NAMED_COLORS: Record<string, string> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  green: '#00FF00',
  blue: '#0000FF',
  cyan: '#00FFFF',
  magenta: '#FF00FF',
  yellow: '#FFFF00'
}

// Excel's 56-color legacy palette (index 1-56). Approximation of the standard palette.
const COLOR_PALETTE: string[] = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
]

function resolveColor(name: string): string | undefined {
  const lower = name.toLowerCase()
  if (lower in NAMED_COLORS) return NAMED_COLORS[lower]
  const m = /^color(\d{1,2})$/i.exec(name)
  if (m) {
    const idx = +m[1]
    if (idx >= 1 && idx <= 56) return COLOR_PALETTE[idx - 1]
  }
  return undefined
}

// ---------------------------------------------------------------------------------------------
// Date/time helpers (1900 date system, including the fake 1900-02-29 leap-year bug)
// ---------------------------------------------------------------------------------------------

interface DateParts {
  y: number
  m: number
  d: number
  dow: number // 0 = Sunday
  hh: number
  mm: number
  ss: number
  ms: number
  secFrac: number // fractional seconds 0..1, full precision
}

function serialToDate(serial: number): DateParts {
  const days = Math.floor(serial)
  const frac = serial - days
  const totalMsRaw = frac * 86400000
  let totalMs = Math.round(totalMsRaw)
  if (totalMs >= 86400000) totalMs -= 86400000
  const hh = Math.floor(totalMs / 3600000)
  const mm = Math.floor((totalMs % 3600000) / 60000)
  const ss = Math.floor((totalMs % 60000) / 1000)
  const ms = totalMs % 1000
  const secFrac = (totalMsRaw / 1000) % 60 - Math.floor((totalMsRaw / 1000) % 60)

  let y: number, m: number, d: number
  if (days === 0) {
    y = 1900; m = 1; d = 0
  } else if (days === 60) {
    y = 1900; m = 2; d = 29
  } else {
    const adj = days > 60 ? days - 1 : days
    const base = Date.UTC(1899, 11, 31) // Sunday
    const dt = new Date(base + adj * 86400000)
    y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate()
  }
  const adjDow = days > 60 ? days - 1 : days
  const dow = ((adjDow % 7) + 7) % 7
  return { y, m, d, dow, hh, mm, ss, ms, secFrac }
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

function pad(n: number, w: number): string {
  const s = String(Math.trunc(n))
  return s.length >= w ? s : '0'.repeat(w - s.length) + s
}

// ---------------------------------------------------------------------------------------------
// Number rounding utilities
// ---------------------------------------------------------------------------------------------

function roundHalfUp(v: number, decimals: number): number {
  if (!isFinite(v)) return v
  const neg = v < 0
  const av = Math.abs(v)
  const factor = Math.pow(10, decimals)
  let scaled = av * factor
  scaled = Number(scaled.toPrecision(15))
  const r = Math.floor(scaled + 0.5)
  const result = r / factor
  return neg ? -result : result
}

// ---------------------------------------------------------------------------------------------
// Format-code parsing
// ---------------------------------------------------------------------------------------------

type Condition = { op: '>' | '<' | '>=' | '<=' | '=' | '<>'; value: number }

interface NumberSection {
  kind: 'number' | 'scientific' | 'fraction' | 'general' | 'literal'
  color?: string
  condition?: Condition
  before: string
  after: string
  intPattern: string[]
  decPattern: string[]
  hasDecimalPoint: boolean
  thousands: boolean
  scale: number
  percent: boolean
  // scientific
  sciExpSign?: '+' | '-'
  sciExpDigits?: number
  // fraction
  fracIntPattern?: string[]
  fracNumDigits?: number
  fracDenDigits?: number
  fracFixedDenom?: number
  fracBefore?: string
  fracAfter?: string
}

type DateToken =
  | { t: 'lit'; s: string }
  | { t: 'y'; n: number }
  | { t: 'd'; n: number }
  | { t: 'h'; n: number }
  | { t: 's'; n: number }
  | { t: 'mon'; n: number } // resolved month
  | { t: 'min'; n: number } // resolved minute
  | { t: 'ampm'; two: boolean }
  | { t: 'elapsed'; unit: 'h' | 'm' | 's'; n: number }
  | { t: 'secfrac'; n: number }

interface DateSection {
  kind: 'date'
  color?: string
  condition?: Condition
  tokens: DateToken[]
  hasAmPm: boolean
}

interface TextSection {
  kind: 'text'
  color?: string
  condition?: Condition
  before: string
  hasAt: boolean
  after: string
}

interface MaskSection {
  kind: 'mask'
  color?: string
  condition?: Condition
  text: string // literal chars, with digit placeholder positions marked '\u0000'
}

type Section = NumberSection | DateSection | TextSection | MaskSection

interface ParsedFormat {
  raw: string
  sections: (Section | undefined)[] // [pos, neg, zero, text]
  numericCount: number // count of numeric (non-text) sections present
  isGeneral: boolean
}

const formatCache = new Map<string, ParsedFormat>()

function evalCond(c: Condition, v: number): boolean {
  switch (c.op) {
    case '>': return v > c.value
    case '<': return v < c.value
    case '>=': return v >= c.value
    case '<=': return v <= c.value
    case '=': return v === c.value
    case '<>': return v !== c.value
  }
}

/** Splits a full format code into up to 4 sections, respecting quotes and brackets. */
function splitSections(code: string): string[] {
  const out: string[] = []
  let cur = ''
  let inQuote = false
  let inBracket = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"' && !inBracket) { inQuote = !inQuote; cur += ch; continue }
    if (!inQuote && ch === '[') { inBracket = true; cur += ch; continue }
    if (!inQuote && ch === ']') { inBracket = false; cur += ch; continue }
    if (ch === ';' && !inQuote && !inBracket) { out.push(cur); cur = ''; continue }
    if (ch === '\\' && !inQuote) { cur += ch; if (i + 1 < code.length) cur += code[++i]; continue }
    cur += ch
  }
  out.push(cur)
  return out
}

/** Extracts leading [condition] and/or [Color] brackets from a section string. */
function extractConditionAndColor(section: string): { rest: string; condition?: Condition; color?: string } {
  let s = section
  let condition: Condition | undefined
  let color: string | undefined
  for (let guard = 0; guard < 2; guard++) {
    const m = /^\[([^\]]*)\](.*)$/.exec(s)
    if (!m) break
    const inner = m[1]
    const condMatch = /^(>=|<=|<>|>|<|=)\s*(-?\d+(?:\.\d+)?)$/.exec(inner)
    if (condMatch) {
      condition = { op: condMatch[1] as Condition['op'], value: +condMatch[2] }
      s = m[2]
      continue
    }
    const col = resolveColor(inner)
    if (col) {
      color = col
      s = m[2]
      continue
    }
    // not a condition/color bracket (e.g. locale/currency or elapsed-time) - stop here, leave in place
    break
  }
  return { rest: s, condition, color }
}

/** Replaces quoted text, backslash-escapes and _x with placeholder markers; drops *x fill chars
 * and resolves inline [$SYMBOL-locale] tags to their literal symbol. Elapsed-time brackets like
 * [h], [mm], [ss] are preserved verbatim for the date tokenizer. */
function preprocess(section: string): { text: string; literals: string[] } {
  const literals: string[] = []
  let out = ''
  let i = 0
  while (i < section.length) {
    const ch = section[i]
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      const lit = end === -1 ? section.slice(i + 1) : section.slice(i + 1, end)
      literals.push(lit)
      out += '\u0000'
      i = end === -1 ? section.length : end + 1
      continue
    }
    if (ch === '\\') {
      const nxt = section[i + 1] ?? ''
      literals.push(nxt)
      out += '\u0000'
      i += 2
      continue
    }
    if (ch === '_') {
      literals.push(' ')
      out += '\u0000'
      i += 2
      continue
    }
    if (ch === '*') {
      const fillCh = section[i + 1] ?? ''
      literals.push(fillCh)
      out += '\u0000'
      i += 2
      continue
    }
    if (ch === '[') {
      const end = section.indexOf(']', i)
      if (end === -1) { out += ch; i++; continue }
      const inner = section.slice(i + 1, end)
      if (/^\$/.test(inner)) {
        const m = /^\$([^-]*)(-.*)?$/.exec(inner)
        const sym = m ? m[1] : ''
        if (sym) { literals.push(sym); out += '\u0000' }
        i = end + 1
        continue
      }
      if (/^[hms]+$/i.test(inner)) {
        out += section.slice(i, end + 1) // keep elapsed-time bracket intact
        i = end + 1
        continue
      }
      // unknown bracket (already-consumed condition/color, or unsupported) - drop
      i = end + 1
      continue
    }
    out += ch
    i++
  }
  return { text: out, literals }
}

function substituteLiterals(text: string, literals: string[]): string {
  let li = 0
  return text.replace(/\u0000/g, () => literals[li++] ?? '')
}

function parseNumberSection(text: string): { intPattern: string[]; decPattern: string[]; hasDecimalPoint: boolean; thousands: boolean; scale: number; before: string; after: string } {
  let i = 0
  let before = ''
  const intPattern: string[] = []
  const decPattern: string[] = []
  let hasDecimalPoint = false
  let thousands = false
  let scale = 0
  let seenDigit = false
  // phase: consume leading literal until first digit-run char
  while (i < text.length && !/[0#?]/.test(text[i])) { before += text[i]; i++ }
  // phase: digit run (digits, '.', ',')
  let seenDecimal = false
  while (i < text.length) {
    const ch = text[i]
    if (ch === '0' || ch === '#' || ch === '?') {
      seenDigit = true
      if (seenDecimal) decPattern.push(ch); else intPattern.push(ch)
      i++
      continue
    }
    if (ch === '.' && !seenDecimal) { seenDecimal = true; hasDecimalPoint = true; i++; continue }
    if (ch === ',') {
      let j = i
      while (text[j] === ',') j++
      const restHasDigit = /[0#?]/.test(text.slice(j))
      if (!restHasDigit && seenDigit) { scale += j - i; i = j; continue }
      thousands = true
      i++
      continue
    }
    break
  }
  const after = text.slice(i)
  return { intPattern, decPattern, hasDecimalPoint, thousands, scale, before, after }
}

function parseFractionSpec(text: string): { before: string; intPattern: string[]; numDigits: number; denDigits: number; fixedDenom?: number; after: string } | null {
  const slashIdx = text.indexOf('/')
  if (slashIdx === -1) return null
  let j = slashIdx - 1
  let num = ''
  while (j >= 0 && /[0#?]/.test(text[j])) { num = text[j] + num; j-- }
  if (!num) return null
  let before = ''
  const intPattern: string[] = []
  if (text[j] === ' ') {
    let k = j - 1
    let ip = ''
    while (k >= 0 && /[0#?]/.test(text[k])) { ip = text[k] + ip; k-- }
    intPattern.push(...ip.split(''))
    before = text.slice(0, k + 1)
  } else {
    before = text.slice(0, j + 1)
  }
  let k2 = slashIdx + 1
  let den = ''
  while (k2 < text.length && /[0#?\d]/.test(text[k2])) { den += text[k2]; k2++ }
  if (!den) return null
  const after = text.slice(k2)
  const fixedDenom = /^\d+$/.test(den) ? parseInt(den, 10) : undefined
  return { before, intPattern, numDigits: num.length, denDigits: den.length, fixedDenom, after }
}

function parseSciSpec(text: string): { before: string; intPattern: string[]; decPattern: string[]; expSign: '+' | '-'; expDigits: number; after: string } | null {
  const m = /^(.*?)([0#]+)(?:\.([0#]+))?[Ee]([+-])(0+)(.*)$/.exec(text)
  if (!m) return null
  return {
    before: m[1],
    intPattern: m[2].split(''),
    decPattern: (m[3] || '').split(''),
    expSign: m[4] as '+' | '-',
    expDigits: m[5].length,
    after: m[6]
  }
}

function parseDateTokens(text: string): DateToken[] {
  const tokens: DateToken[] = []
  let i = 0
  while (i < text.length) {
    const rest = text.slice(i)
    let m: RegExpExecArray | null
    if ((m = /^am\/pm/i.exec(rest))) { tokens.push({ t: 'ampm', two: false }); i += 5; continue }
    if ((m = /^a\/p/i.exec(rest))) { tokens.push({ t: 'ampm', two: true }); i += 3; continue }
    if ((m = /^\[(h+|m+|s+)\]/i.exec(rest))) {
      const unit = m[1][0].toLowerCase() as 'h' | 'm' | 's'
      tokens.push({ t: 'elapsed', unit, n: m[1].length })
      i += m[0].length
      continue
    }
    const ch = text[i]
    if (/[ymdhsYMDHS]/.test(ch)) {
      let n = 1
      while (text[i + n] && text[i + n].toLowerCase() === ch.toLowerCase()) n++
      const lower = ch.toLowerCase()
      if (lower === 'y') tokens.push({ t: 'y', n })
      else if (lower === 'd') tokens.push({ t: 'd', n })
      else if (lower === 'h') tokens.push({ t: 'h', n })
      else if (lower === 's') tokens.push({ t: 's', n })
      else tokens.push({ t: 'mon', n }) // 'm' placeholder, resolved later
      i += n
      continue
    }
    if (ch === '\u0000') { tokens.push({ t: 'lit', s: '\u0000' }); i++; continue }
    // accumulate literal run
    let s = ''
    while (i < text.length && !/[ymdhsYMDHS\u0000]/.test(text[i]) && !/^(am\/pm|a\/p)/i.test(text.slice(i)) && !/^\[(h+|m+|s+)\]/i.test(text.slice(i))) {
      s += text[i]; i++
    }
    tokens.push({ t: 'lit', s })
  }
  // resolve secfrac: an 's' token immediately followed by a lit "." + zeros
  for (let k = 0; k < tokens.length - 1; k++) {
    const a = tokens[k], b = tokens[k + 1]
    if (a.t === 's' && b.t === 'lit') {
      const fm = /^\.(0+)/.exec(b.s)
      if (fm) {
        tokens.splice(k + 1, 1, { t: 'secfrac', n: fm[1].length }, { t: 'lit', s: b.s.slice(fm[0].length) })
      }
    }
  }
  // resolve 'mon' -> minute when adjacent to h/elapsed-h before, or s/elapsed-s/secfrac after
  for (let k = 0; k < tokens.length; k++) {
    if (tokens[k].t !== 'mon') continue
    let isMinute = false
    for (let p = k - 1; p >= 0; p--) {
      const tk = tokens[p]
      if (tk.t === 'lit') continue
      if (tk.t === 'h' || (tk.t === 'elapsed' && tk.unit === 'h')) isMinute = true
      break
    }
    if (!isMinute) {
      for (let n = k + 1; n < tokens.length; n++) {
        const tk = tokens[n]
        if (tk.t === 'lit') continue
        if (tk.t === 's' || tk.t === 'secfrac' || (tk.t === 'elapsed' && tk.unit === 's')) isMinute = true
        break
      }
    }
    const cur = tokens[k] as { t: 'mon'; n: number }
    tokens[k] = isMinute ? { t: 'min', n: cur.n } : { t: 'mon', n: cur.n }
  }
  return tokens
}

function parseSection(raw: string): Section {
  const { rest, condition, color } = extractConditionAndColor(raw)
  const { text, literals } = preprocess(rest)

  if (/@/.test(text)) {
    const idx = text.indexOf('@')
    const beforeRaw = text.slice(0, idx)
    const afterRaw = text.slice(idx + 1)
    const beforeCount = countPlaceholders(beforeRaw)
    const afterCount = countPlaceholders(afterRaw)
    return {
      kind: 'text',
      color, condition,
      before: substituteLiterals(beforeRaw, literals.slice(0, beforeCount)),
      hasAt: true,
      after: substituteLiterals(afterRaw, literals.slice(beforeCount, beforeCount + afterCount))
    }
  }

  const isDate = /(?:^|[^A-Za-z])(?:am\/pm|a\/p)/i.test(text) || /[ymdhsYMDHS]/.test(text.replace(/\u0000/g, ''))
  if (isDate) {
    const tokens = parseDateTokens(text)
    // substitute literal placeholders within lit tokens
    let li = 0
    for (const tk of tokens) if (tk.t === 'lit') tk.s = tk.s.replace(/\u0000/g, () => literals[li++] ?? '')
    const hasAmPm = tokens.some((t) => t.t === 'ampm')
    return { kind: 'date', color, condition, tokens, hasAmPm }
  }

  const percent = text.includes('%')

  const sci = parseSciSpec(text)
  if (sci) {
    const before = substituteLiterals(sci.before, sliceLiterals(sci.before, literals))
    const after = substituteLiterals(sci.after, sliceLiterals(sci.after, literals, sci.before))
    return {
      kind: 'scientific', color, condition,
      before,
      after,
      intPattern: sci.intPattern,
      decPattern: sci.decPattern,
      hasDecimalPoint: sci.decPattern.length > 0,
      thousands: false, scale: 0, percent,
      sciExpSign: sci.expSign, sciExpDigits: sci.expDigits
    }
  }

  const frac = parseFractionSpec(text)
  if (frac) {
    return {
      kind: 'fraction', color, condition,
      before: substituteLiterals(frac.before, sliceLiterals(frac.before, literals)),
      after: substituteLiterals(frac.after, sliceLiterals(frac.after, literals, frac.before)),
      intPattern: [], decPattern: [], hasDecimalPoint: false, thousands: false, scale: 0, percent: false,
      fracIntPattern: frac.intPattern,
      fracNumDigits: frac.numDigits,
      fracDenDigits: frac.denDigits,
      fracFixedDenom: frac.fixedDenom
    }
  }

  if (text.replace(/[0#?.,%\u0000]/g, '') === text && !/[0#?]/.test(text) && text.trim() === '') {
    // no placeholders at all - literal only (rare)
  }

  const firstRun = parseNumberSection(text)
  if (/[0#?]/.test(firstRun.after) && !firstRun.thousands && firstRun.scale === 0 && !firstRun.hasDecimalPoint) {
    // multiple digit-placeholder runs separated by literal text (zip+4, phone, SSN masks)
    let li = 0
    let outFinal = ''
    for (const ch of text) {
      if (ch === '0' || ch === '#' || ch === '?') outFinal += '\u0001' // digit slot marker
      else if (ch === '\u0000') outFinal += literals[li++] ?? '' // literal placeholder
      else outFinal += ch
    }
    return { kind: 'mask', color, condition, text: outFinal }
  }

  const num = firstRun
  const before = substituteLiterals(num.before, sliceLiterals(num.before, literals))
  const after = substituteLiterals(num.after, sliceLiterals(num.after, literals, num.before + num.intPattern.join('') + (num.hasDecimalPoint ? '.' : '') + num.decPattern.join('') + ','.repeat(num.scale)))
  return {
    kind: 'number', color, condition,
    before, after,
    intPattern: num.intPattern, decPattern: num.decPattern,
    hasDecimalPoint: num.hasDecimalPoint, thousands: num.thousands, scale: num.scale, percent
  }
}

// helpers to slice the literals array by counting placeholder markers consumed so far
function countPlaceholders(s: string): number {
  return (s.match(/\u0000/g) || []).length
}
function sliceLiterals(segment: string, all: string[], precedingConsumed?: string): string[] {
  const start = precedingConsumed ? countPlaceholders(precedingConsumed) : 0
  const n = countPlaceholders(segment)
  return all.slice(start, start + n)
}

function parseFormat(code: string): ParsedFormat {
  const cached = formatCache.get(code)
  if (cached) return cached

  let parsed: ParsedFormat
  if (code === '' || code === 'General' || code == null) {
    parsed = { raw: code, sections: [undefined, undefined, undefined, undefined], numericCount: 1, isGeneral: true }
  } else {
    const rawSections = splitSections(code)
    const sections: (Section | undefined)[] = [undefined, undefined, undefined, undefined]
    for (let i = 0; i < Math.min(rawSections.length, 4); i++) {
      const s = rawSections[i].trim() === '' ? undefined : parseSection(rawSections[i])
      sections[i] = s
    }
    const numericCount = sections.slice(0, 3).filter((s) => s && s.kind !== 'text').length || (rawSections.length === 1 ? 1 : 0)
    parsed = { raw: code, sections, numericCount: numericCount || 1, isGeneral: false }
  }
  formatCache.set(code, parsed)
  return parsed
}

// ---------------------------------------------------------------------------------------------
// General format rendering
// ---------------------------------------------------------------------------------------------

function toExcelSci(av: number, decimals = 5): string {
  if (av === 0) return (0).toFixed(decimals) + 'E+00'
  let exp = Math.floor(Math.log10(av))
  let mant = av / Math.pow(10, exp)
  let mantRounded = Number(mant.toFixed(decimals))
  if (mantRounded >= 10) { mantRounded /= 10; exp++ }
  if (mantRounded < 1) { mantRounded *= 10; exp-- }
  const mantStr = mantRounded.toFixed(decimals)
  const expSign = exp >= 0 ? '+' : '-'
  const expAbs = Math.abs(exp)
  const expStr = (expAbs < 10 ? '0' : '') + expAbs
  return mantStr + 'E' + expSign + expStr
}

function formatGeneral(value: number): string {
  if (value === 0) return '0'
  const neg = value < 0
  const av = Math.abs(value)
  let s: string
  if (av >= 1e11 || av < 1e-9) {
    s = toExcelSci(av, 5)
  } else {
    s = av.toPrecision(11)
    if (s.includes('e')) {
      s = toExcelSci(av, 5)
    } else {
      if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '')
    }
  }
  return neg ? '-' + s : s
}

// ---------------------------------------------------------------------------------------------
// Number section rendering
// ---------------------------------------------------------------------------------------------

function renderIntCore(pattern: string[], intDigitsStr: string): string {
  const n = Math.max(pattern.length, intDigitsStr.length)
  let out = ''
  const digitsPadded = intDigitsStr.padStart(n, ' ')
  const patPadded = Array(n - pattern.length).fill('#').concat(pattern)
  for (let i = 0; i < n; i++) {
    const d = digitsPadded[i]
    const p = patPadded[i]
    if (d !== ' ') out += d
    else if (p === '0') out += '0'
    else if (p === '?') out += ' '
  }
  return out
}

function renderIntWithGrouping(pattern: string[], intDigitsStr: string, thousands: boolean): string {
  const core = renderIntCore(pattern, intDigitsStr)
  if (!thousands) return core
  let i = 0
  while (i < core.length && core[i] === ' ') i++
  const lead = core.slice(0, i)
  const digits = core.slice(i)
  if (!digits) return core
  const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return lead + grouped
}

function renderDecCore(pattern: string[], fracStr: string): string {
  let end = fracStr.length
  while (end > 0 && pattern[end - 1] === '#' && fracStr[end - 1] === '0') end--
  return fracStr.slice(0, end)
}

function renderNumber(sec: NumberSection, av: number): string {
  if (sec.intPattern.length === 0 && sec.decPattern.length === 0 && !sec.hasDecimalPoint) {
    return sec.before + sec.after
  }
  let v = av
  if (sec.percent) v *= 100
  if (sec.scale) v /= Math.pow(1000, sec.scale)
  const decimals = sec.decPattern.length
  const rounded = roundHalfUp(v, decimals)
  const fixed = rounded.toFixed(decimals)
  const [intStr, fracStr = ''] = fixed.split('.')
  const intAbs = intStr.replace('-', '')
  let intDigitsStr = intAbs
  if (intAbs === '0' && !sec.intPattern.includes('0')) intDigitsStr = ''
  const intRendered = renderIntWithGrouping(sec.intPattern, intDigitsStr, sec.thousands)
  const decRendered = decimals > 0 ? renderDecCore(sec.decPattern, fracStr) : ''
  let core: string
  if (intRendered === '' && decRendered === '' && !sec.intPattern.includes('0')) {
    core = ''
  } else if (decimals > 0 && decRendered !== '') {
    core = intRendered + '.' + decRendered
  } else if (sec.hasDecimalPoint && decRendered === '' && sec.decPattern.some((p) => p === '0')) {
    core = intRendered + '.' + fracStr
  } else {
    core = intRendered
  }
  return sec.before + core + sec.after
}

function renderScientific(sec: NumberSection, av: number): string {
  const decimals = sec.decPattern.length
  const k = Math.max(1, sec.intPattern.length)
  let exp = 0
  let mantInt = 0
  let fracStr = ''
  if (av !== 0) {
    const rawExp = Math.floor(Math.log10(av))
    let adjExp = Math.floor(rawExp / k) * k
    let rounded = roundHalfUp(av / Math.pow(10, adjExp), decimals)
    if (rounded >= Math.pow(10, k)) { adjExp += k; rounded = roundHalfUp(av / Math.pow(10, adjExp), decimals) }
    exp = adjExp
    const fixed = rounded.toFixed(decimals)
    const [i, f = ''] = fixed.split('.')
    mantInt = +i
    fracStr = f
  } else {
    fracStr = '0'.repeat(decimals)
  }
  const intRendered = renderIntCore(sec.intPattern.length ? sec.intPattern : ['0'], String(mantInt))
  const expSign = exp >= 0 ? '+' : '-'
  const expDigits = sec.sciExpDigits ?? 2
  const expStr = pad(Math.abs(exp), expDigits)
  const core = intRendered + (decimals > 0 ? '.' + fracStr : '') + 'E' + expSign + expStr
  return sec.before + core + sec.after
}

function gcd(a: number, b: number): number { return b === 0 ? a : gcd(b, a % b) }

function renderFraction(sec: NumberSection, av: number): string {
  const hasInt = (sec.fracIntPattern?.length ?? 0) > 0
  let intPart = 0
  let frac = av
  if (hasInt) {
    intPart = Math.floor(av)
    frac = av - intPart
  }
  let num: number, den: number
  if (sec.fracFixedDenom) {
    den = sec.fracFixedDenom
    num = Math.round(frac * den)
    if (num === den) { num = 0; intPart += 1 }
  } else {
    const maxDen = Math.pow(10, sec.fracDenDigits ?? 1) - 1
    const best = bestFraction(frac, maxDen)
    num = best.num; den = best.den
    if (den === 0) { num = 0; den = 1 }
    if (num === den) { num = 0; den = 1; intPart += 1 }
  }
  if (num === 0 && hasInt) {
    const intStr = renderIntCore(sec.fracIntPattern!, String(intPart))
    return sec.before + intStr + sec.after
  }
  const numOut = String(num)
  const denOut = sec.fracFixedDenom ? String(sec.fracFixedDenom) : String(den)
  const body = numOut + '/' + denOut
  const prefix = hasInt ? renderIntCore(sec.fracIntPattern!, intPart === 0 ? '' : String(intPart)) + ' ' : ''
  return sec.before + prefix + body + sec.after
}

function bestFraction(x: number, maxDen: number): { num: number; den: number } {
  if (x === 0) return { num: 0, den: 1 }
  let bestNum = 0, bestDen = 1, bestErr = Math.abs(x)
  for (let den = 1; den <= maxDen; den++) {
    const num = Math.round(x * den)
    const err = Math.abs(x - num / den)
    if (err < bestErr - 1e-12) { bestErr = err; bestNum = num; bestDen = den }
  }
  const g = gcd(bestNum, bestDen) || 1
  return { num: bestNum / g, den: bestDen / g }
}

// ---------------------------------------------------------------------------------------------
// Date section rendering
// ---------------------------------------------------------------------------------------------

function renderDate(sec: DateSection, serial: number): string {
  const dp = serialToDate(serial)
  const hour12 = sec.hasAmPm ? (dp.hh % 12 === 0 ? 12 : dp.hh % 12) : dp.hh
  const totalMs = serial * 86400000
  let out = ''
  for (const tk of sec.tokens) {
    switch (tk.t) {
      case 'lit': out += tk.s; break
      case 'y': out += tk.n <= 2 ? pad(dp.y % 100, 2) : pad(dp.y, 4); break
      case 'd':
        if (tk.n === 1) out += String(dp.d)
        else if (tk.n === 2) out += pad(dp.d, 2)
        else if (tk.n === 3) out += DAY_NAMES[dp.dow].slice(0, 3)
        else out += DAY_NAMES[dp.dow]
        break
      case 'mon':
        if (tk.n === 1) out += String(dp.m)
        else if (tk.n === 2) out += pad(dp.m, 2)
        else if (tk.n === 3) out += MONTH_NAMES[dp.m - 1].slice(0, 3)
        else if (tk.n === 4) out += MONTH_NAMES[dp.m - 1]
        else out += MONTH_NAMES[dp.m - 1].slice(0, 1)
        break
      case 'min':
        out += tk.n === 1 ? String(dp.mm) : pad(dp.mm, 2)
        break
      case 'h':
        out += tk.n === 1 ? String(hour12) : pad(hour12, 2)
        break
      case 's':
        out += tk.n === 1 ? String(dp.ss) : pad(dp.ss, 2)
        break
      case 'secfrac': {
        const frac = Math.round(dp.secFrac * Math.pow(10, tk.n))
        out += '.' + pad(frac, tk.n)
        break
      }
      case 'ampm':
        out += dp.hh >= 12 ? (tk.two ? 'P' : 'PM') : (tk.two ? 'A' : 'AM')
        break
      case 'elapsed': {
        let total: number
        if (tk.unit === 'h') total = Math.floor(totalMs / 3600000)
        else if (tk.unit === 'm') total = Math.floor(totalMs / 60000)
        else total = Math.floor(totalMs / 1000)
        out += pad(total, tk.n)
        break
      }
    }
  }
  return out
}

// ---------------------------------------------------------------------------------------------
// Text section rendering
// ---------------------------------------------------------------------------------------------

function renderMask(sec: MaskSection, av: number): string {
  const slots = (sec.text.match(/\u0001/g) || []).length
  const digits = String(Math.round(av)).padStart(slots, '0')
  let di = 0
  let out = ''
  for (const ch of sec.text) out += ch === '\u0001' ? digits[di++] ?? '0' : ch
  return out
}

function renderTextSection(sec: TextSection, value: string): FormattedValue {
  const text = sec.before + value + sec.after
  return { text, color: sec.color, align: 'left' }
}

// ---------------------------------------------------------------------------------------------
// Section selection
// ---------------------------------------------------------------------------------------------

function pickSection(parsed: ParsedFormat, value: number): { sec: Section | undefined; isNegSection: boolean } {
  const secs = parsed.sections
  const numeric = [secs[0], secs[1], secs[2]]
  for (const s of numeric) {
    if (s && 'condition' in s && s.condition && evalCond(s.condition, value)) {
      return { sec: s, isNegSection: false }
    }
  }
  const count = parsed.numericCount
  if (count <= 1) return { sec: secs[0], isNegSection: false }
  if (count === 2) return value < 0 ? { sec: secs[1], isNegSection: true } : { sec: secs[0], isNegSection: false }
  // 3+
  if (value > 0) return { sec: secs[0], isNegSection: false }
  if (value < 0) return { sec: secs[1] ?? secs[0], isNegSection: true }
  return { sec: secs[2] ?? secs[0], isNegSection: false }
}

// ---------------------------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------------------------

/** Formats a computed value with an Excel format code ('General' when undefined). */
export function formatValue(value: CellValue, numFmt?: string, opts?: { colWidthPx?: number }): FormattedValue {
  void opts
  if (value === null || value === undefined) return { text: '' }
  if (isError(value)) return { text: value.error, align: 'center' }
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', align: 'center' }

  const parsed = parseFormat(numFmt && numFmt.length ? numFmt : 'General')

  if (typeof value === 'string') {
    if (parsed.isGeneral) return { text: value, align: 'left' }
    const definedCount = parsed.sections.filter(Boolean).length
    let sec = parsed.sections[3]
    if (!sec && definedCount === 1 && parsed.sections[0]?.kind === 'text') sec = parsed.sections[0]
    if (sec && sec.kind === 'text') return renderTextSection(sec, value)
    return { text: value, align: 'left' }
  }

  // numeric
  if (parsed.isGeneral) {
    return { text: formatGeneral(value), align: 'right' }
  }

  const { sec } = pickSection(parsed, value)
  if (!sec || sec.kind === 'text') {
    return { text: formatGeneral(value), align: 'right' }
  }

  const av = Math.abs(value)
  let text: string
  switch (sec.kind) {
    case 'date': text = renderDate(sec, value); break
    case 'fraction': text = renderFraction(sec, av); break
    case 'scientific': text = renderScientific(sec, av); break
    case 'mask': text = renderMask(sec, av); break
    case 'general': text = formatGeneral(value); break
    default: text = renderNumber(sec, av); break
  }
  const autoSign = parsed.numericCount <= 1 && value < 0 && sec.kind !== 'date' && /[1-9]/.test(text)
  if (autoSign) text = '-' + text
  return { text, color: (sec as { color?: string }).color, align: 'right' }
}

// ---------------------------------------------------------------------------------------------
// detectInputFormat
// ---------------------------------------------------------------------------------------------

/**
 * Called when the user types into a cell. If the text implies a format (e.g. '$1,200', '15%',
 * '3/4/2025', '1.5E+3'), returns the format code Excel would auto-apply; otherwise undefined.
 */
export function detectInputFormat(raw: string): string | undefined {
  const s = raw.trim()
  if (!s) return undefined

  // Currency: $1,234.50 / $5
  let m = /^\$\s*-?[\d,]+(\.\d+)?$/.exec(s)
  if (m) {
    const decMatch = /\.(\d+)$/.exec(s)
    return decMatch ? '"$"#,##0.' + '0'.repeat(decMatch[1].length) : '$#,##0'
  }

  // Percent: 15% / 12.5%
  m = /^-?\d+(\.\d+)?%$/.exec(s)
  if (m) {
    return m[1] ? '0.00%' : '0%'
  }

  // Scientific: 1.5E+3
  m = /^-?\d+(\.\d+)?[eE][+-]?\d+$/.exec(s)
  if (m) return '0.00E+00'

  // Time: 10:30, 10:30 PM, 10:30:15
  m = /^(\d{1,2}):(\d{2})(:(\d{2}))?\s*(AM|PM|am|pm)?$/.exec(s)
  if (m) {
    const hasSec = !!m[3]
    const hasAmPm = !!m[5]
    if (hasAmPm) return hasSec ? 'h:mm:ss AM/PM' : 'h:mm AM/PM'
    return hasSec ? 'h:mm:ss' : 'h:mm'
  }

  // Date: 1/31/2025, 1/31, 31-Jan, Jan 2025
  m = /^(\d{1,2})([/.-])(\d{1,2})\2(\d{2,4})$/.exec(s)
  if (m) return SHORT_DATE_FORMAT
  m = /^(\d{1,2})\/(\d{1,2})$/.exec(s)
  if (m) return 'd-mmm'
  m = /^(\d{1,2})-([A-Za-z]{3,9})$/.exec(s)
  if (m && MONTH_NAMES.some((mn) => mn.toLowerCase().startsWith(m![2].toLowerCase()))) return 'd-mmm'
  m = /^([A-Za-z]{3,9})\s+(\d{4})$/.exec(s)
  if (m && MONTH_NAMES.some((mn) => mn.toLowerCase().startsWith(m![1].toLowerCase()))) return 'mmm-yy'
  m = /^([A-Za-z]{3,9})-(\d{2,4})$/.exec(s)
  if (m && MONTH_NAMES.some((mn) => mn.toLowerCase().startsWith(m![1].toLowerCase()))) return 'mmm-yy'

  // Thousands: 1,234 / 1,234.5
  m = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/.exec(s)
  if (m) return m[2] ? '#,##0.00' : '#,##0'

  return undefined
}

// ---------------------------------------------------------------------------------------------
// FORMAT_CATEGORIES (Format Cells > Number tab data)
// ---------------------------------------------------------------------------------------------

export interface FormatEntry {
  label: string // as shown in the Format Cells dialog sample list
  code: string
}

export const FORMAT_CATEGORIES: Record<string, FormatEntry[]> = {
  General: [{ label: 'General', code: 'General' }],
  Number: [
    { label: '1234', code: '0' },
    { label: '1234.00', code: '0.00' },
    { label: '1,234', code: '#,##0' },
    { label: '1,234.00', code: '#,##0.00' },
    { label: '1234 (red negative)', code: '0;[Red]-0' }
  ],
  Currency: [
    { label: '$1,234.00', code: '$#,##0.00' },
    { label: '$1,234', code: '$#,##0_);($#,##0)' },
    { label: '$1,234.00 (red negative)', code: '$#,##0.00_);[Red]($#,##0.00)' },
    { label: '$1,234', code: '$#,##0;[Red]$#,##0' }
  ],
  Accounting: [
    { label: '$ 1,234.00 ', code: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
    { label: ' 1,234.00 ', code: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)' }
  ],
  Date: [
    { label: '3/14/2012', code: 'm/d/yyyy' },
    { label: 'Wednesday, March 14, 2012', code: 'dddd, mmmm d, yyyy' },
    { label: '3/14', code: 'm/d' },
    { label: '3/14/12', code: 'm/d/yy' },
    { label: '03/14/12', code: 'mm/dd/yy' },
    { label: '14-Mar', code: 'd-mmm' },
    { label: '14-Mar-12', code: 'd-mmm-yy' },
    { label: 'Mar-12', code: 'mmm-yy' },
    { label: 'March-12', code: 'mmmm-yy' },
    { label: 'March 14, 2012', code: 'mmmm d, yyyy' },
    { label: '3/14/12 1:30 PM', code: 'm/d/yy h:mm AM/PM' },
    { label: '3/14/12 13:30', code: 'm/d/yy h:mm' },
    { label: 'M', code: 'M' },
    { label: 'M-12', code: 'M-yy' },
    { label: '14-Mar-2012', code: 'd-mmm-yyyy' }
  ],
  Time: [
    { label: '1:30 PM', code: 'h:mm AM/PM' },
    { label: '1:30:55 PM', code: 'h:mm:ss AM/PM' },
    { label: '13:30', code: 'h:mm' },
    { label: '13:30:55', code: 'h:mm:ss' },
    { label: '37:30:55', code: '[h]:mm:ss' },
    { label: '30:55.0', code: 'mm:ss.0' }
  ],
  Percentage: [
    { label: '15%', code: '0%' },
    { label: '15.25%', code: '0.00%' }
  ],
  Fraction: [
    { label: '1/2', code: '# ?/?' },
    { label: '1/50', code: '# ??/??' },
    { label: '1/2 (eighths)', code: '# ?/8' }
  ],
  Scientific: [
    { label: '1.23E+05', code: '0.00E+00' },
    { label: '12.3E+04', code: '##0.0E+0' }
  ],
  Text: [{ label: '@', code: '@' }],
  Special: [
    { label: 'Zip Code', code: '00000' },
    { label: 'Zip Code + 4', code: '00000-0000' },
    { label: 'Phone Number', code: '[<=9999999]###-####;(###) ###-####' },
    { label: 'Social Security Number', code: '000-00-0000' }
  ],
  Custom: [
    { label: 'General', code: 'General' },
    { label: '0', code: '0' },
    { label: '0.00', code: '0.00' },
    { label: '#,##0', code: '#,##0' },
    { label: '#,##0.00', code: '#,##0.00' },
    { label: '#,##0_);(#,##0)', code: '#,##0_);(#,##0)' },
    { label: '#,##0_);[Red](#,##0)', code: '#,##0_);[Red](#,##0)' },
    { label: '#,##0.00_);(#,##0.00)', code: '#,##0.00_);(#,##0.00)' },
    { label: '#,##0.00_);[Red](#,##0.00)', code: '#,##0.00_);[Red](#,##0.00)' },
    { label: '$#,##0_);($#,##0)', code: '$#,##0_);($#,##0)' },
    { label: '$#,##0_);[Red]($#,##0)', code: '$#,##0_);[Red]($#,##0)' },
    { label: '$#,##0.00_);($#,##0.00)', code: '$#,##0.00_);($#,##0.00)' },
    { label: '$#,##0.00_);[Red]($#,##0.00)', code: '$#,##0.00_);[Red]($#,##0.00)' },
    { label: '0%', code: '0%' },
    { label: '0.00%', code: '0.00%' },
    { label: '0.00E+00', code: '0.00E+00' },
    { label: '# ?/?', code: '# ?/?' },
    { label: '# ??/??', code: '# ??/??' },
    { label: 'm/d/yyyy', code: 'm/d/yyyy' },
    { label: 'd-mmm-yy', code: 'd-mmm-yy' },
    { label: 'd-mmm', code: 'd-mmm' },
    { label: 'mmm-yy', code: 'mmm-yy' },
    { label: 'h:mm AM/PM', code: 'h:mm AM/PM' },
    { label: 'h:mm:ss AM/PM', code: 'h:mm:ss AM/PM' },
    { label: 'h:mm', code: 'h:mm' },
    { label: 'h:mm:ss', code: 'h:mm:ss' },
    { label: 'm/d/yyyy h:mm', code: 'm/d/yyyy h:mm' },
    { label: 'mm:ss', code: 'mm:ss' },
    { label: '[h]:mm:ss', code: '[h]:mm:ss' },
    { label: 'mm:ss.0', code: 'mm:ss.0' },
    { label: '##0.0E+0', code: '##0.0E+0' },
    { label: '@', code: '@' },
    { label: 'Long Date', code: '[$-x-sysdate]dddd, mmmm d, yyyy' },
    { label: 'Long Time', code: '[$-x-systime]h:mm:ss AM/PM' }
  ]
}

// ---------------------------------------------------------------------------------------------
// Format-builder helpers (Format Cells dialog / Increase-Decrease Decimal buttons)
// ---------------------------------------------------------------------------------------------

export function buildNumberFormat(opts: {
  decimals: number
  thousands: boolean
  negative: 0 | 1 | 2 | 3 // 0: -1,234.10  1: (1,234.10) red  2: -1,234.10 red  3: (1,234.10)
  symbol?: string // e.g. '$' for currency; undefined for plain number
}): string {
  const dec = opts.decimals > 0 ? '.' + '0'.repeat(opts.decimals) : ''
  const int = opts.thousands ? '#,##0' : '0'
  const sym = opts.symbol ? '"' + opts.symbol + '"' : ''
  const pos = sym + int + dec
  switch (opts.negative) {
    case 0: return pos + ';-' + pos
    case 1: return pos + ';(' + pos + ')'
    case 2: return pos + ';[Red]-' + pos
    case 3: return pos + ';[Red](' + pos + ')'
    default: return pos
  }
}

/** Increase (+1) or decrease (-1) the number of decimal places shown by a format code.
 * Works on plain numeric formats and on 'General' (using the value's currently-displayed
 * decimal count as the starting point). */
export function adjustDecimals(code: string, delta: 1 | -1, value?: number): string {
  const isGeneral = !code || code === 'General'
  if (isGeneral) {
    const shown = value !== undefined ? formatGeneral(value) : '0'
    const decMatch = /\.(\d+)$/.exec(shown)
    const startDecimals = decMatch ? decMatch[1].length : 0
    const next = Math.max(0, startDecimals + delta)
    return next === 0 ? '0' : '0.' + '0'.repeat(next)
  }
  // find first run of decimal-digit placeholders (0/#/?) after a '.', outside brackets/quotes
  const parts = splitSections(code)
  const rebuilt = parts.map((section) => adjustSectionDecimals(section, delta))
  return rebuilt.join(';')
}

function adjustSectionDecimals(section: string, delta: 1 | -1): string {
  // locate a '.' followed by digit placeholders, not inside [] or ""
  let inQuote = false, inBracket = false
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') { inQuote = !inQuote; continue }
    if (ch === '[') inBracket = true
    if (ch === ']') inBracket = false
    if (inQuote || inBracket) continue
    if (ch === '.') {
      let j = i + 1
      let run = ''
      while (j < section.length && /[0#?]/.test(section[j])) { run += section[j]; j++ }
      if (delta === 1) {
        const added = run.length ? run[run.length - 1] : '0'
        return section.slice(0, j) + added + section.slice(j)
      } else {
        if (run.length <= 1) {
          // remove the whole decimal point + last digit
          return section.slice(0, i) + section.slice(j)
        }
        return section.slice(0, j - 1) + section.slice(j)
      }
    }
  }
  // no decimal point found: only add on increase, inserting after the last integer digit placeholder
  if (delta === 1) {
    let lastDigit = -1
    inQuote = false; inBracket = false
    for (let i = 0; i < section.length; i++) {
      const ch = section[i]
      if (ch === '"') { inQuote = !inQuote; continue }
      if (ch === '[') inBracket = true
      if (ch === ']') inBracket = false
      if (inQuote || inBracket) continue
      if (/[0#?]/.test(ch)) lastDigit = i
    }
    if (lastDigit === -1) return section
    return section.slice(0, lastDigit + 1) + '.0' + section.slice(lastDigit + 1)
  }
  return section
}

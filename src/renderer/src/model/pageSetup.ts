// Page Layout / Page Setup: paper sizes, margins and the header & footer code language of Excel.
import { PageSetup } from './types'

export interface PaperSize { code: number; name: string; wIn: number; hIn: number }

/** Excel's ST_PaperSize codes (the subset that appears in Excel's own Size menu). */
export const PAPER_SIZES: PaperSize[] = [
  { code: 1, name: 'Letter', wIn: 8.5, hIn: 11 },
  { code: 3, name: 'Tabloid', wIn: 11, hIn: 17 },
  { code: 5, name: 'Legal', wIn: 8.5, hIn: 14 },
  { code: 6, name: 'Statement', wIn: 5.5, hIn: 8.5 },
  { code: 7, name: 'Executive', wIn: 7.25, hIn: 10.5 },
  { code: 8, name: 'A3', wIn: 11.69, hIn: 16.54 },
  { code: 9, name: 'A4', wIn: 8.27, hIn: 11.69 },
  { code: 11, name: 'A5', wIn: 5.83, hIn: 8.27 },
  { code: 12, name: 'B4 (JIS)', wIn: 10.12, hIn: 14.33 },
  { code: 13, name: 'B5 (JIS)', wIn: 7.17, hIn: 10.12 },
  { code: 14, name: 'Folio', wIn: 8.5, hIn: 13 }
]
export const DEFAULT_PAPER = 9 // A4, like Excel on an English (India) PC

export function paperOf(ps: PageSetup | undefined): PaperSize {
  return PAPER_SIZES.find((p) => p.code === (ps?.paperSize ?? DEFAULT_PAPER)) ?? PAPER_SIZES.find((p) => p.code === DEFAULT_PAPER)!
}

export interface Margins { left: number; right: number; top: number; bottom: number; header: number; footer: number }
export const MARGIN_PRESETS: Record<'normal' | 'wide' | 'narrow', Margins> = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
}
export function marginsOf(ps: PageSetup | undefined): Margins {
  return { ...MARGIN_PRESETS.normal, ...(ps?.margins ?? {}) }
}
export function marginPresetOf(ps: PageSetup | undefined): 'normal' | 'wide' | 'narrow' | 'custom' {
  const m = marginsOf(ps)
  for (const k of ['normal', 'wide', 'narrow'] as const) {
    const p = MARGIN_PRESETS[k]
    if ((Object.keys(p) as (keyof Margins)[]).every((f) => Math.abs(p[f] - m[f]) < 0.005)) return k
  }
  return 'custom'
}

/** Paper size in inches after orientation. */
export function pageInches(ps: PageSetup | undefined): { w: number; h: number } {
  const p = paperOf(ps)
  return ps?.landscape ? { w: p.hIn, h: p.wIn } : { w: p.wIn, h: p.hIn }
}

/** Scale to Fit settings as Excel names them: fitToPage with width/height page counts (0 = automatic). */
export function fitOf(ps: PageSetup | undefined): { on: boolean; w: number; h: number } {
  return { on: !!ps?.fitToPage, w: ps?.fitToWidth ?? 1, h: ps?.fitToHeight ?? 1 }
}

// ---------- header & footer codes ----------
export type HfSection = 'left' | 'center' | 'right'
export interface HfParts { left: string; center: string; right: string }

/** Splits "&Lleft&Ccenter&Rright" into its three sections (codes inside are kept verbatim). */
export function splitHf(code: string | undefined): HfParts {
  const out: HfParts = { left: '', center: '', right: '' }
  if (!code) return out
  let cur: HfSection = 'center'
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '&' && i + 1 < code.length) {
      const n = code[i + 1]
      if (n === 'L' || n === 'C' || n === 'R') { cur = n === 'L' ? 'left' : n === 'C' ? 'center' : 'right'; i++; continue }
      out[cur] += ch + n; i++; continue
    }
    out[cur] += ch
  }
  return out
}
export function joinHf(p: HfParts): string {
  return (p.left ? '&L' + p.left : '') + (p.center ? '&C' + p.center : '') + (p.right ? '&R' + p.right : '')
}

export interface HfContext { page: number; pages: number; sheet: string; file: string; path: string; now: Date }

export interface HfRun { text: string; bold?: boolean; italic?: boolean; underline?: 'single' | 'double'; strike?: boolean; sup?: boolean; sub?: boolean; font?: string; size?: number; color?: string }

/** Expands one section's codes (&P &N &D &T &F &A &Z, fonts, sizes, colors...) into styled text runs. */
export function hfRuns(section: string, ctx: HfContext): HfRun[] {
  const runs: HfRun[] = []
  let st: Omit<HfRun, 'text'> = {}
  let buf = ''
  const flush = (): void => { if (buf) { runs.push({ text: buf, ...st }); buf = '' } }
  const dateText = ctx.now.toLocaleDateString('en-GB')
  const timeText = ctx.now.toLocaleTimeString('en-US')
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch !== '&') { buf += ch; continue }
    const n = section[i + 1]
    if (n === undefined) break
    i++
    if (n === '&') buf += '&'
    else if (n === 'P') buf += ctx.page
    else if (n === 'N') buf += ctx.pages
    else if (n === 'D') buf += dateText
    else if (n === 'T') buf += timeText
    else if (n === 'F') buf += ctx.file
    else if (n === 'A') buf += ctx.sheet
    else if (n === 'Z') buf += ctx.path
    else if (n === 'G') { /* picture: not supported */ }
    else if (n === 'B') { flush(); st = { ...st, bold: !st.bold } }
    else if (n === 'I') { flush(); st = { ...st, italic: !st.italic } }
    else if (n === 'U') { flush(); st = { ...st, underline: st.underline === 'single' ? undefined : 'single' } }
    else if (n === 'E') { flush(); st = { ...st, underline: st.underline === 'double' ? undefined : 'double' } }
    else if (n === 'S') { flush(); st = { ...st, strike: !st.strike } }
    else if (n === 'X') { flush(); st = { ...st, sup: !st.sup } }
    else if (n === 'Y') { flush(); st = { ...st, sub: !st.sub } }
    else if (n === '"') {
      const end = section.indexOf('"', i + 1)
      const spec = section.slice(i + 1, end === -1 ? section.length : end)
      i = end === -1 ? section.length : end
      flush()
      const [font, style] = spec.split(',')
      st = { ...st, font: font || undefined }
      if (style) { const s = style.toLowerCase(); st = { ...st, bold: /bold/.test(s), italic: /italic/.test(s) } }
    } else if (/[0-9]/.test(n)) {
      let num = n
      while (/[0-9]/.test(section[i + 1] ?? '')) num += section[++i]
      flush(); st = { ...st, size: parseInt(num, 10) }
    } else if (n === 'K') {
      const hex = section.slice(i + 1, i + 7)
      if (/^[0-9A-Fa-f]{6}$/.test(hex)) { flush(); st = { ...st, color: '#' + hex }; i += 6 }
    } else buf += '&' + n
  }
  flush()
  return runs
}

/** Section text with the codes stripped, e.g. for a one-line preview. */
export function hfPlain(section: string, ctx: HfContext): string {
  return hfRuns(section, ctx).map((r) => r.text).join('')
}

/** Excel's built-in Header / Footer dropdown choices (each is a full &L&C&R code string). */
export function hfPresets(ctx: { user: string; company?: string }): { label: string; code: string }[] {
  return [
    { label: '(none)', code: '' },
    { label: 'Page 1', code: ('&CPage &P') },
    { label: 'Page 1 of ?', code: ('&CPage &P of &N') },
    { label: 'Sheet1', code: ('&C&A') },
    { label: 'Confidential, dd-mm-yyyy, Page 1', code: ('&LConfidential&C&D&RPage &P') },
    { label: 'Book1', code: ('&C&F') },
    { label: 'Sheet1, Page 1', code: ('&C&A, Page &P') },
    { label: 'Book1, Page 1', code: ('&C&F, Page &P') },
    { label: 'Sheet1, Book1, Page 1', code: ('&L&A&C&F&RPage &P') },
    { label: 'Prepared by ' + ctx.user + ' dd-mm-yyyy, Page 1', code: (`&LPrepared by ${ctx.user} &D&RPage &P`) }
  ]
}

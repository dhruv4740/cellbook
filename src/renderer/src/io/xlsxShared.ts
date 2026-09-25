// Small, dependency-free XLSX helpers shared by the heavy exceljs-based reader/writer (xlsx.ts)
// and the fast fflate-based reader (fastXlsx.ts / xlsx.worker.ts). Keeping these free of exceljs,
// HyperFormula and React lets the worker bundle stay small.
import type { BorderLineStyle, CommentReply, CommentThread, FilterCriteria, NoteData, Range } from '../model/types'
import { BLANK_TOKEN } from '../engine/filter'
import { decodeXmlEntities, extractSection, parseAttrs } from './fastXlsxXml'

// ---------- column width / row height conversions ----------
// Excel character-width <-> pixel conversion for the default 11pt font (Maximum Digit Width ~7px).
export function colWidthToPx(w: number): number {
  return Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7)
}
export function pxToColWidth(px: number): number {
  if (px <= 0) return 0
  return Math.round(((px * 256) / 7 - Math.trunc(128 / 7)) / 256 * 100) / 100
}
// Row height: points <-> pixels at 96 dpi.
export function ptToPx(pt: number): number {
  return Math.round((pt * 4) / 3)
}
export function pxToPt(px: number): number {
  return Math.round(((px * 3) / 4) * 100) / 100
}

// ---------- Office default theme colors (Office 2013+ theme) ----------
export const THEME_COLORS = [
  'FFFFFF', // 0 bg1 (light1)
  '000000', // 1 tx1 (dark1)
  'E7E6E6', // 2 bg2 (light2)
  '44546A', // 3 tx2 (dark2)
  '4472C4', // 4 accent1
  'ED7D31', // 5 accent2
  'A5A5A5', // 6 accent3
  'FFC000', // 7 accent4
  '5B9BD5', // 8 accent5
  '70AD47' // 9 accent6
]

export function applyTint(hex: string, tint: number): string {
  // Excel tint: positive lightens toward white, negative darkens toward black.
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  const adj = (c: number): number => {
    if (tint < 0) return Math.round(c * (1 + tint))
    return Math.round(c * (1 - tint) + 255 * tint)
  }
  const toHex = (n: number): string => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')
  return (toHex(adj(r)) + toHex(adj(g)) + toHex(adj(b))).toUpperCase()
}

// ---------- borders ----------
export const BORDER_STYLES: BorderLineStyle[] = [
  'thin', 'medium', 'thick', 'double', 'dotted', 'dashed', 'hair',
  'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'
]

// ---------- range refs ----------
export function colLetterToIndex(s: string): number {
  let n = 0
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}
export function parseRangeStr(s: string): Range | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(s)
  if (!m) return null
  return { r1: +m[2] - 1, c1: colLetterToIndex(m[1]), r2: +m[4] - 1, c2: colLetterToIndex(m[3]) }
}

// ---------- AutoFilter (read side) ----------
// exceljs doesn't expose per-column filter criteria at all, and the fast reader's worksheet
// scanner only ever captured the filter range - both readers need the same filterColumn parsing,
// so it lives here instead of being duplicated in fastXlsx.ts and xlsx.ts.
export function parseAutoFilterRef(xml: string): string | undefined {
  const m = /<autoFilter ref="([^"]+)"/.exec(xml)
  return m ? m[1] : undefined
}

/** Per-column checked-value criteria from `<autoFilter>...<filterColumn>...</autoFilter>`, keyed
 *  by colId *relative to the filter range's first column* (per the OOXML spec) - callers add the
 *  range's c1 to get an absolute column index, per FilterCriteria's own key convention. */
export function parseAutoFilterColumns(sheetXml: string): Record<number, FilterCriteria> {
  const columns: Record<number, FilterCriteria> = {}
  const section = extractSection(sheetXml, 'autoFilter')
  if (!section) return columns
  const colRe = /<filterColumn\b([^>]*)>([\s\S]*?)<\/filterColumn>/g
  let m: RegExpExecArray | null
  while ((m = colRe.exec(section))) {
    const attrs = parseAttrs(m[1], 0, m[1].length)
    const colId = attrs.colId !== undefined ? parseInt(attrs.colId, 10) : NaN
    if (!Number.isFinite(colId)) continue
    const body = m[2]
    const filtersM = /<filters\b([^>]*?)\/?>/.exec(body)
    if (!filtersM) continue
    const filtersAttrs = parseAttrs(filtersM[1], 0, filtersM[1].length)
    const selected: string[] = []
    if (filtersAttrs.blank === '1') selected.push(BLANK_TOKEN)
    const valRe = /<filter val="([^"]*)"/g
    let vm: RegExpExecArray | null
    while ((vm = valRe.exec(body))) selected.push(decodeXmlEntities(vm[1]))
    columns[colId] = { selected }
  }
  return columns
}

// ---------- threaded-comment legacy fallback (read + write side) ----------
// A real threaded comment lives in a separate XML part newer Excel versions understand; for
// compatibility Excel 2021 *also* writes an ordinary legacy note (xl/comments*.xml) whose text is
// this fixed fallback message, and reconstructs the thread from that same text on open. We do the
// same so a cell's `comments` thread survives a save/reopen instead of silently vanishing, without
// having to emit the (much more involved) threadedComment part.
const THREAD_PREFIX = '[Threaded comment]'
const THREAD_BOILERPLATE =
  '\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    '
const REPLY_MARKER = '\nReply:\n    '

export function threadToLegacyNote(thread: CommentThread): NoteData {
  let text = THREAD_PREFIX + THREAD_BOILERPLATE + thread.text
  for (const reply of thread.replies) text += REPLY_MARKER + reply.text
  return { author: thread.author, text }
}

/** Reverses threadToLegacyNote. Reply authors/times aren't recoverable from the fallback text
 *  (real Excel doesn't preserve them there either), so replies default to the thread's own author. */
export function legacyNoteToThread(note: NoteData): CommentThread | null {
  if (!note.text.startsWith(THREAD_PREFIX)) return null
  const marker = 'Comment:\n    '
  const idx = note.text.indexOf(marker)
  const rest = idx === -1 ? '' : note.text.slice(idx + marker.length)
  const parts = rest.split(REPLY_MARKER)
  const replies: CommentReply[] = parts.slice(1).map((text, i) => ({ id: 'r' + i, author: note.author, text, time: '' }))
  return { author: note.author, text: parts[0] ?? '', time: '', replies }
}

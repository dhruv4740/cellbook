// Sheet / workbook protection: Excel's password hashes and the allow-list defaults.
import { sha512 } from '@noble/hashes/sha2.js'
import { EditRange, PasswordHash, ProtectAllowKey, Range, SheetProtection, parseA1, rangeA1 } from './types'

export const PROTECT_OPTIONS: { key: ProtectAllowKey; label: string }[] = [
  { key: 'selectLockedCells', label: 'Select locked cells' },
  { key: 'selectUnlockedCells', label: 'Select unlocked cells' },
  { key: 'formatCells', label: 'Format cells' },
  { key: 'formatColumns', label: 'Format columns' },
  { key: 'formatRows', label: 'Format rows' },
  { key: 'insertColumns', label: 'Insert columns' },
  { key: 'insertRows', label: 'Insert rows' },
  { key: 'insertHyperlinks', label: 'Insert hyperlinks' },
  { key: 'deleteColumns', label: 'Delete columns' },
  { key: 'deleteRows', label: 'Delete rows' },
  { key: 'sort', label: 'Sort' },
  { key: 'autoFilter', label: 'Use AutoFilter' },
  { key: 'pivotTables', label: 'Use PivotTable and PivotChart' },
  { key: 'objects', label: 'Edit objects' },
  { key: 'scenarios', label: 'Edit scenarios' }
]

/** What Excel's Protect Sheet dialog starts with: only selecting cells is allowed. */
export function defaultAllow(): SheetProtection['allow'] {
  const a = {} as SheetProtection['allow']
  for (const o of PROTECT_OPTIONS) a[o.key] = o.key === 'selectLockedCells' || o.key === 'selectUnlockedCells'
  return a
}

const enc = new TextEncoder()
function b64(bytes: Uint8Array): string {
  let s = ''
  for (const b of bytes) s += String.fromCharCode(b)
  return btoa(s)
}
function unb64(s: string): Uint8Array {
  const bin = atob(s)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}
function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2)
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); out[i * 2] = c & 255; out[i * 2 + 1] = c >> 8 }
  return out
}
function agile(password: string, salt: Uint8Array, spin: number): string {
  const pw = utf16le(password)
  const first = new Uint8Array(salt.length + pw.length)
  first.set(salt); first.set(pw, salt.length)
  let h = sha512(first)
  const buf = new Uint8Array(h.length + 4)
  const dv = new DataView(buf.buffer)
  for (let i = 0; i < spin; i++) {
    buf.set(h); dv.setUint32(h.length, i, true)
    h = sha512(buf)
  }
  return b64(h)
}

/** ECMA-376 legacy 16-bit password hash (the `password="CC1A"` attribute of older files). */
export function legacyHash(password: string): string {
  let hash = 0
  const n = password.length
  for (let i = 0; i < n; i++) {
    let c = password.charCodeAt(i) & 0x7fff
    c = ((c << (i + 1)) & 0x7fff) | (c >> (15 - (i + 1)))
    hash ^= c
  }
  hash ^= n ^ 0xce4b
  return hash.toString(16).toUpperCase().padStart(4, '0')
}

export function hashPassword(password: string): PasswordHash | undefined {
  if (!password) return undefined
  const salt = crypto.getRandomValues(new Uint8Array(16))
  return { hashValue: agile(password, salt, 100000), saltValue: b64(salt), spinCount: 100000 }
}

export function checkPassword(hash: PasswordHash | undefined, password: string): boolean {
  if (!hash) return true
  if (hash.legacy !== undefined) return legacyHash(password) === hash.legacy.toUpperCase()
  if (!hash.hashValue || !hash.saltValue) return true
  return agile(password, unb64(hash.saltValue), hash.spinCount ?? 100000) === hash.hashValue
}

/** XML attribute fragment for a password hash (leading space included), or ''. */
export function passwordAttrs(h: PasswordHash | undefined, prefix: '' | 'workbook'): string {
  if (!h) return ''
  const k = (n: string) => (prefix ? prefix + n[0].toUpperCase() + n.slice(1) : n)
  if (h.legacy !== undefined) return ` ${prefix ? 'workbookPassword' : 'password'}="${h.legacy}"`
  if (!h.hashValue) return ''
  return ` ${k('algorithmName')}="SHA-512" ${k('hashValue')}="${h.hashValue}" ${k('saltValue')}="${h.saltValue}" ${k('spinCount')}="${h.spinCount ?? 100000}"`
}

// Attributes that are prohibitions when "1" (attribute absent → the default below applies).
const PROHIBIT_DEFAULT: Record<ProtectAllowKey, boolean> = {
  selectLockedCells: false, selectUnlockedCells: false, formatCells: true, formatColumns: true, formatRows: true,
  insertColumns: true, insertRows: true, insertHyperlinks: true, deleteColumns: true, deleteRows: true,
  sort: true, autoFilter: true, pivotTables: true, objects: false, scenarios: false
}

export function parseSheetProtection(attrs: Record<string, string>): SheetProtection | undefined {
  if (attrs.sheet !== '1' && attrs.sheet !== 'true') return undefined
  const allow = {} as SheetProtection['allow']
  for (const o of PROTECT_OPTIONS) {
    const v = attrs[o.key]
    const prohibited = v === undefined ? PROHIBIT_DEFAULT[o.key] : v === '1' || v === 'true'
    allow[o.key] = !prohibited
  }
  const out: SheetProtection = { allow }
  if (attrs.hashValue && attrs.saltValue) out.password = { hashValue: attrs.hashValue, saltValue: attrs.saltValue, spinCount: parseInt(attrs.spinCount ?? '100000', 10) }
  else if (attrs.password) out.password = { legacy: attrs.password }
  return out
}

export function sheetProtectionXml(p: SheetProtection): string {
  let a = passwordAttrs(p.password, '') + ' sheet="1"'
  // Excel's attribute order: objects and scenarios come right after sheet.
  const ordered = [...PROTECT_OPTIONS.filter((o) => o.key === 'objects' || o.key === 'scenarios'), ...PROTECT_OPTIONS.filter((o) => o.key !== 'objects' && o.key !== 'scenarios')]
  for (const o of ordered) {
    const prohibited = !p.allow[o.key]
    // Like Excel, write only attributes that differ from the schema default.
    if (prohibited !== PROHIBIT_DEFAULT[o.key]) a += ` ${o.key}="${prohibited ? 1 : 0}"`
  }
  return `<sheetProtection${a}/>`
}

/** The ranges named by an sqref-style string ("A1:B4 D6"); unparsable parts are dropped. */
export function parseSqref(ref: string): Range[] {
  return ref.split(/[\s,]+/).filter(Boolean).map((p) => parseA1(p)).filter((r): r is Range => !!r)
}

export function inEditRanges(list: EditRange[] | undefined, r: number, c: number): boolean {
  return !!list?.some((e) => parseSqref(e.ref).some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2))
}

const escAttr = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** The sheet's <protectedRanges> element, or '' when there are none. */
export function protectedRangesXml(list: EditRange[] | undefined): string {
  const ok = (list ?? []).filter((e) => parseSqref(e.ref).length)
  if (!ok.length) return ''
  return `<protectedRanges>${ok.map((e) => `<protectedRange sqref="${escAttr(parseSqref(e.ref).map(rangeA1).join(' '))}" name="${escAttr(e.title)}"/>`).join('')}</protectedRanges>`
}

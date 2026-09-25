// Reads a legacy Excel 97-2003 .xls (BIFF8, inside an OLE2/CFB compound file) into the same
// WorkbookData shape the .xlsx readers produce (see io/fastXlsx.ts). Kept free of
// exceljs/hyperformula/react so it can run inside io/xlsx.worker.ts.
//
// Record-layout notes throughout this file are cross-checked against Apache POI's HSSF record
// classes (org.apache.poi.hssf.record.*), which are the most reliable public reference for BIFF8's
// undocumented-in-practice bit packing (colors, borders, alignment).
import * as CFB from 'cfb'
import type { BorderLineStyle, BorderSide, CellStyle, NoteData, Range, SheetLayout } from '../../model/types'
import type { WorkbookData } from '../../model/store'
import { defaultLayout, DEFAULT_AUTHOR } from '../../model/defaults'
import { THEME_COLORS as _THEME, BORDER_STYLES, colWidthToPx, ptToPx } from '../xlsxShared'
import { BIFF_ERR, BUILTIN_NUMFMTS, REC } from './biffConst'
import { decodeFormula, type ExternSheetRange, type FormulaCtx } from './formula'

void _THEME // xlsxShared re-export not needed here; keep import list minimal/explicit otherwise.

// ---------- low-level record framing ----------
interface Rec {
  type: number
  data: Uint8Array
  bounds: number[] // offsets (within `data`) where a CONTINUE segment begins; used for
  // boundary-aware string reads (a string split across a CONTINUE needs a fresh option-flags byte).
  offset: number // this record's type+len header offset in the Workbook stream (matches BOUNDSHEET.lbPlyPos)
}

function readRecords(buf: Uint8Array): Rec[] {
  const out: Rec[] = []
  const n = buf.length
  let pos = 0
  while (pos + 4 <= n) {
    const offset = pos
    const type = buf[pos] | (buf[pos + 1] << 8)
    const len = buf[pos + 2] | (buf[pos + 3] << 8)
    pos += 4
    if (pos + len > n) break // truncated - stop gracefully rather than throw on a corrupt tail
    const first = buf.subarray(pos, pos + len)
    pos += len
    if (type === REC.CONTINUE) continue // stray CONTINUE with no owner - ignore

    const segs = [first]
    const bounds: number[] = []
    let total = first.length
    while (pos + 4 <= n) {
      const t2 = buf[pos] | (buf[pos + 1] << 8)
      if (t2 !== REC.CONTINUE) break
      const l2 = buf[pos + 2] | (buf[pos + 3] << 8)
      const p2 = pos + 4
      if (p2 + l2 > n) break
      const seg = buf.subarray(p2, p2 + l2)
      bounds.push(total)
      segs.push(seg)
      total += seg.length
      pos = p2 + l2
    }
    let data: Uint8Array
    if (segs.length === 1) data = first
    else {
      data = new Uint8Array(total)
      let o = 0
      for (const s of segs) { data.set(s, o); o += s.length }
    }
    out.push({ type, data, bounds, offset })
  }
  return out
}

// ---------- byte/string reader over one (possibly CONTINUE-merged) record's data ----------
class Rdr {
  pos = 0
  constructor(public data: Uint8Array, public bounds: number[] = []) {}
  get left(): number { return this.data.length - this.pos }
  u8(): number { return this.data[this.pos++] }
  u16(): number { const v = this.data[this.pos] | (this.data[this.pos + 1] << 8); this.pos += 2; return v }
  i16(): number { const v = this.u16(); return v >= 0x8000 ? v - 0x10000 : v }
  u32(): number {
    const v = (this.data[this.pos] | (this.data[this.pos + 1] << 8) | (this.data[this.pos + 2] << 16) | (this.data[this.pos + 3] << 24)) >>> 0
    this.pos += 4
    return v
  }
  f64(): number {
    const dv = new DataView(this.data.buffer, this.data.byteOffset + this.pos, 8)
    this.pos += 8
    return dv.getFloat64(0, true)
  }
  bytes(n: number): Uint8Array { const b = this.data.subarray(this.pos, this.pos + n); this.pos += n; return b }
  rest(): Uint8Array { return this.bytes(this.left) }
  skip(n: number): void { this.pos += n }

  /** BIFF8 (rich/ext) unicode string, CONTINUE-boundary aware: per spec, if a string's character
   *  data is split by a CONTINUE, the continuation starts with a fresh 1-byte option-flags byte. */
  readStr(cchWidth: 1 | 2, allowRichExt: boolean): string {
    const cch = cchWidth === 1 ? this.u8() : this.u16()
    let grbit = this.u8()
    let wide = (grbit & 1) !== 0
    let cRun = 0
    let cbExtRst = 0
    if (allowRichExt) {
      if (grbit & 0x08) cRun = this.u16()
      if (grbit & 0x04) cbExtRst = this.u32()
    }
    let s = ''
    let remaining = cch
    while (remaining > 0) {
      let nextBoundary = Infinity
      for (const b of this.bounds) if (b > this.pos && b < nextBoundary) nextBoundary = b
      const avail = nextBoundary === Infinity ? this.left : nextBoundary - this.pos
      const charBytes = wide ? 2 : 1
      const canRead = Math.max(0, Math.min(remaining, Math.floor(avail / charBytes)))
      for (let i = 0; i < canRead; i++) s += String.fromCharCode(wide ? this.u16() : this.u8())
      remaining -= canRead
      if (remaining > 0) {
        // Landed exactly on a CONTINUE boundary with characters left: re-read the option flags.
        grbit = this.u8()
        wide = (grbit & 1) !== 0
      }
    }
    if (allowRichExt) {
      if (cRun) this.skip(cRun * 4)
      if (cbExtRst) this.skip(cbExtRst)
    }
    return s
  }

  /** Plain (no option-flags byte) character read for a known length and known width - used by
   *  NAME, which reads its own multibyte flag earlier in the record instead of a per-string one. */
  readStrKnownLenPlain(cch: number, wide: boolean): string {
    let s = ''
    for (let i = 0; i < cch; i++) s += String.fromCharCode(wide ? this.u16() : this.u8())
    return s
  }

  /** Same boundary-aware character read as readStr, but for a known character count with no
   *  length prefix (used for TXO's comment text, which carries its length in the TXO header). */
  readStrKnownLen(cch: number): string {
    let grbit = this.u8()
    let wide = (grbit & 1) !== 0
    let s = ''
    let remaining = cch
    while (remaining > 0) {
      let nextBoundary = Infinity
      for (const b of this.bounds) if (b > this.pos && b < nextBoundary) nextBoundary = b
      const avail = nextBoundary === Infinity ? this.left : nextBoundary - this.pos
      const charBytes = wide ? 2 : 1
      const canRead = Math.max(0, Math.min(remaining, Math.floor(avail / charBytes)))
      for (let i = 0; i < canRead; i++) s += String.fromCharCode(wide ? this.u16() : this.u8())
      remaining -= canRead
      if (remaining > 0) { grbit = this.u8(); wide = (grbit & 1) !== 0 }
    }
    return s
  }
}

// ---------- RK / cell-number decoding ----------
function decodeRK(rk: number): number {
  const isInt = (rk & 2) !== 0
  const fx100 = (rk & 1) !== 0
  let num: number
  if (isInt) {
    num = rk >> 2 // arithmetic shift keeps the sign
  } else {
    const buf = new ArrayBuffer(8)
    const dv = new DataView(buf)
    dv.setUint32(0, 0, true)
    dv.setUint32(4, rk & 0xfffffffc, true)
    num = dv.getFloat64(0, true)
  }
  return fx100 ? num / 100 : num
}

// ---------- number formatting (dates) ----------
function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false
  const stripped = fmt.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')
  return /[ymdhs]/i.test(stripped) && !/^general$/i.test(fmt)
}
const DATE_1904_OFFSET = 1462 // days between the 1900 and 1904 date systems' epochs

// ---------- default 56-color BIFF palette (used when the file has no PALETTE record) ----------
const DEFAULT_PALETTE: string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
]
const FIXED8 = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF']
function icvColor(icv: number, palette: string[]): string | undefined {
  if (icv < 8) return FIXED8[icv]
  if (icv >= 8 && icv <= 63) return '#' + (palette[icv - 8] ?? DEFAULT_PALETTE[icv - 8] ?? '000000')
  return undefined // 64/65 = automatic text/background, everything else = unmapped system color
}

const BORDER_STYLE_BY_ID: (BorderLineStyle | undefined)[] = [
  undefined, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair',
  'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'
]
function borderSide(styleId: number, icv: number, palette: string[]): BorderSide | undefined {
  const style = BORDER_STYLE_BY_ID[styleId]
  if (!style || !(BORDER_STYLES as readonly string[]).includes(style)) return undefined
  return { style, color: icvColor(icv, palette) }
}

const HALIGN = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'] as const
const VALIGN = ['top', 'middle', 'bottom', 'justify', 'distributed'] as const

// ---------- FONT ----------
interface FontRaw { name: string; size: number; bold: boolean; italic: boolean; strike: boolean; underline: 0 | 1 | 2; icv: number }
function parseFont(r: Rdr): FontRaw {
  const dyHeight = r.u16()
  const attrs = r.u16()
  const italic = (attrs & 0x02) !== 0
  const strike = (attrs & 0x08) !== 0
  const icv = r.u16()
  const bls = r.u16()
  r.u16() // super/subscript - not modeled by CellStyle
  const uls = r.u8()
  r.skip(3) // family, charset, reserved
  const name = r.readStr(1, false)
  const underline: 0 | 1 | 2 = uls === 1 || uls === 0x21 ? 1 : uls === 2 || uls === 0x22 ? 2 : 0
  return { name, size: dyHeight / 20, bold: bls >= 700, italic, strike, underline, icv }
}
function fontToStyle(f: FontRaw, palette: string[]): CellStyle {
  const s: CellStyle = {}
  if (f.name && f.name !== 'Calibri') s.fontFamily = f.name
  if (f.size && f.size !== 11) s.fontSize = f.size
  if (f.bold) s.bold = true
  if (f.italic) s.italic = true
  if (f.strike) s.strike = true
  if (f.underline === 1) s.underline = 'single'
  else if (f.underline === 2) s.underline = 'double'
  if (f.icv !== 0x7fff) {
    const c = icvColor(f.icv, palette)
    if (c) s.color = c
  }
  return s
}

// ---------- XF ----------
interface XfRaw {
  ifnt: number; ifmt: number; locked: boolean; hidden: boolean
  alc: number; wrap: boolean; alcV: number; rotation: number
  indent: number; shrinkToFit: boolean
  borderLeft: number; borderRight: number; borderTop: number; borderBottom: number
  icvLeft: number; icvRight: number; icvTop: number; icvBottom: number
  fillPattern: number; icvFore: number; icvBack: number
}
function parseXf(r: Rdr): XfRaw {
  const ifnt = r.u16()
  const ifmt = r.u16()
  const cellOptions = r.u16()
  const locked = (cellOptions & 0x0001) !== 0
  const hidden = (cellOptions & 0x0002) !== 0
  const alignOpts = r.u16()
  const alc = alignOpts & 0x07
  const wrap = (alignOpts & 0x08) !== 0
  const alcV = (alignOpts >> 4) & 0x07
  const trot = (alignOpts >> 8) & 0xff
  const indentOpts = r.u16()
  const indent = indentOpts & 0x0f
  const shrinkToFit = (indentOpts & 0x10) !== 0
  const borderOpts = r.u16()
  const borderLeft = borderOpts & 0x0f
  const borderRight = (borderOpts >> 4) & 0x0f
  const borderTop = (borderOpts >> 8) & 0x0f
  const borderBottom = (borderOpts >> 12) & 0x0f
  const paletteOpts = r.u16()
  const icvLeft = paletteOpts & 0x7f
  const icvRight = (paletteOpts >> 7) & 0x7f
  const adtl = r.u32()
  const icvTop = adtl & 0x7f
  const icvBottom = (adtl >> 7) & 0x7f
  const fillPattern = (adtl >>> 26) & 0x3f
  const fillOpts = r.u16()
  const icvFore = fillOpts & 0x7f
  const icvBack = (fillOpts >> 7) & 0x7f
  const rotation = trot === 255 ? 255 : trot <= 90 ? trot : -(trot - 90)
  return { ifnt, ifmt, locked, hidden, alc, wrap, alcV, rotation, indent, shrinkToFit, borderLeft, borderRight, borderTop, borderBottom, icvLeft, icvRight, icvTop, icvBottom, fillPattern, icvFore, icvBack }
}
function xfToStyle(xf: XfRaw, fonts: FontRaw[], formats: Map<number, string>, palette: string[]): CellStyle | undefined {
  const s: CellStyle = {}
  const fontId = xf.ifnt < 4 ? xf.ifnt : xf.ifnt - 1 // font index 4 doesn't exist (BIFF quirk)
  const font = fonts[fontId]
  if (font) Object.assign(s, fontToStyle(font, palette))
  const code = formats.get(xf.ifmt) ?? BUILTIN_NUMFMTS[xf.ifmt]
  if (code && code !== 'General') s.numFmt = code
  if (xf.alc !== 0) s.hAlign = HALIGN[xf.alc]
  if (xf.alcV !== 2) s.vAlign = VALIGN[xf.alcV]
  if (xf.wrap) s.wrap = true
  if (xf.shrinkToFit) s.shrinkToFit = true
  if (xf.indent) s.indent = xf.indent
  if (xf.rotation) s.rotation = xf.rotation
  const borders: NonNullable<CellStyle['borders']> = {}
  const l = borderSide(xf.borderLeft, xf.icvLeft, palette); if (l) borders.left = l
  const rr = borderSide(xf.borderRight, xf.icvRight, palette); if (rr) borders.right = rr
  const t = borderSide(xf.borderTop, xf.icvTop, palette); if (t) borders.top = t
  const b = borderSide(xf.borderBottom, xf.icvBottom, palette); if (b) borders.bottom = b
  if (Object.keys(borders).length) s.borders = borders
  if (xf.fillPattern === 1) {
    const c = icvColor(xf.icvFore, palette)
    if (c) s.fill = c
  }
  if (!xf.locked) s.locked = false
  if (xf.hidden) s.hidden = true
  return Object.keys(s).length ? s : undefined
}

// ---------- formula cached value ----------
interface CachedFormulaValue { kind: 'num' | 'string' | 'bool' | 'error' | 'empty'; value: number | string | boolean }
function readFormulaCachedValue(r: Rdr): CachedFormulaValue {
  const raw = r.bytes(8)
  const dv = new DataView(raw.buffer, raw.byteOffset, 8)
  if (dv.getUint16(6, true) === 0xffff) {
    const type = raw[0]
    const data = raw[2]
    if (type === 0) return { kind: 'string', value: '' } // actual text is in the following STRING record
    if (type === 1) return { kind: 'bool', value: data !== 0 }
    if (type === 2) return { kind: 'error', value: BIFF_ERR[data] ?? '#N/A' }
    return { kind: 'empty', value: '' }
  }
  return { kind: 'num', value: dv.getFloat64(0, true) }
}

// ---------- workbook globals ----------
interface Globals {
  fonts: FontRaw[]
  formats: Map<number, string>
  xfStyles: (CellStyle | undefined)[]
  palette: string[]
  dateMode1904: boolean
  boundSheets: { name: string; lbPlyPos: number; sheetType: number; hidden: number }[]
  sst: string[]
  externSheets: ExternSheetRange[]
  names: string[]
}

function parseGlobals(records: Rec[], globalsEnd: number): Globals {
  const fonts: FontRaw[] = []
  const formats = new Map<number, string>()
  const xfRaw: XfRaw[] = []
  let palette: string[] = []
  let dateMode1904 = false
  const boundSheets: Globals['boundSheets'] = []
  let sst: string[] = []
  const supbooks: { internal: boolean }[] = []
  const externSheetRaw: { iSupBook: number; first: number; last: number }[] = []
  const names: string[] = ['']

  for (let i = 0; i < globalsEnd; i++) {
    const rec = records[i]
    const r = new Rdr(rec.data, rec.bounds)
    switch (rec.type) {
      case REC.FILEPASS:
        throw new Error('This workbook is password-protected; Excel Clone cannot open encrypted .xls files.')
      case REC.FONT:
        fonts.push(parseFont(r))
        break
      case REC.FORMAT: {
        const ifmt = r.u16()
        formats.set(ifmt, r.readStr(2, false))
        break
      }
      case REC.XF:
        xfRaw.push(parseXf(r))
        break
      case REC.PALETTE: {
        const count = r.u16()
        const cols: string[] = []
        for (let k = 0; k < count; k++) {
          const red = r.u8(); const green = r.u8(); const blue = r.u8(); r.u8()
          cols.push([red, green, blue].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase())
        }
        palette = cols
        break
      }
      case REC.DATEMODE:
        dateMode1904 = r.u16() === 1
        break
      case REC.BOUNDSHEET: {
        const lbPlyPos = r.u32()
        const grbit = r.u16()
        const name = r.readStr(1, false)
        boundSheets.push({ name, lbPlyPos, sheetType: (grbit >> 8) & 0xff, hidden: grbit & 0x03 })
        break
      }
      case REC.SST: {
        r.u32() // total (non-unique) count, unused here
        const unique = r.u32()
        const strs: string[] = []
        for (let k = 0; k < unique; k++) {
          try { strs.push(r.readStr(2, true)) } catch { strs.push('') }
        }
        sst = strs
        break
      }
      case REC.SUPBOOK: {
        const numSheets = r.u16()
        void numSheets
        if (rec.data.length <= 4) {
          const tag = r.u16()
          supbooks.push({ internal: tag === 0x0401 })
        } else {
          supbooks.push({ internal: false })
        }
        break
      }
      case REC.EXTERNSHEET: {
        const count = r.u16()
        for (let k = 0; k < count; k++) {
          const iSupBook = r.u16(); const first = r.u16(); const last = r.u16()
          externSheetRaw.push({ iSupBook, first, last })
        }
        break
      }
      case REC.NAME: {
        // Fixed 15-byte header before the name text; only the display text is needed here (for
        // ptgName), not the name's own definition/formula.
        try {
          const optionFlag = r.u16()
          r.u8() // keyboard shortcut
          const cch = r.u8()
          r.u16() // length of name definition (rgce) - not decoded
          r.u16() // externSheetIndex+1
          r.u16() // sheetNumber
          r.u8(); r.u8(); r.u8(); r.u8() // custom menu / description / help / status-bar text lengths
          const isMultibyte = r.u8() !== 0
          const isBuiltin = (optionFlag & 0x20) !== 0
          if (isBuiltin) {
            r.u8()
            names.push('')
          } else {
            names.push(r.readStrKnownLenPlain(cch, isMultibyte))
          }
        } catch {
          names.push('')
        }
        break
      }
      default:
        break
    }
  }

  const xfStyles = xfRaw.map((xf) => xfToStyle(xf, fonts, formats, palette))
  const externSheets: ExternSheetRange[] = externSheetRaw.map(({ iSupBook, first, last }) => {
    const sb = supbooks[iSupBook]
    return sb?.internal ? { first, last } : { first: -1, last: -1 }
  })

  return { fonts, formats, xfStyles, palette, dateMode1904, boundSheets, sst, externSheets, names }
}

// ---------- per-sheet parsing ----------
interface SharedGroup { r1: number; c1: number; r2: number; c2: number; rgce: Uint8Array; extra?: Uint8Array; isArray: boolean }

function inGroup(g: SharedGroup, r: number, c: number): boolean {
  return r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2
}

function parseSheet(
  records: Rec[], start: number, end: number, sheetNames: string[], names: string[], externSheets: ExternSheetRange[],
  xfStyles: (CellStyle | undefined)[], dateMode1904: boolean, sst: string[]
): { cells: WorkbookData['sheets'][number]['cells']; layout: Partial<SheetLayout> } {
  const cells: WorkbookData['sheets'][number]['cells'] = Object.create(null)
  const colWidths: Record<number, number> = {}
  const hiddenCols: Record<number, true> = {}
  const rowHeights: Record<number, number> = {}
  const hiddenRows: Record<number, true> = {}
  const merges: Range[] = []
  const notes: Record<string, NoteData> = {}
  let defColWidthChars: number | undefined
  let showGridlines = true
  let showHeadings = true
  let zoom = 100
  let frozen: { rows: number; cols: number } | undefined
  let freezePanesFlag = false
  let paneSplit: { x: number; y: number } | undefined

  const groups: SharedGroup[] = []
  const pendingNoteTexts: string[] = []
  let lastNoteAnchor: { row: number; col: number } | undefined

  const setCell = (r: number, c: number, raw: string | undefined, style: CellStyle | undefined): void => {
    if (raw !== undefined || style) cells[r + ',' + c] = { raw, style }
  }

  const styleOf = (ixfe: number): CellStyle | undefined => xfStyles[ixfe]

  const ctxFor = (row: number, col: number): FormulaCtx => ({ baseRow: row, baseCol: col, sheetNames, externSheets, names })

  const maybeDateAdjust = (n: number, style: CellStyle | undefined): number => {
    if (dateMode1904 && style?.numFmt && isDateFormat(style.numFmt)) return n + DATE_1904_OFFSET
    return n
  }

  for (let i = start; i < end; i++) {
    const rec = records[i]
    const r = new Rdr(rec.data, rec.bounds)
    switch (rec.type) {
      case REC.DEFCOLWIDTH:
        defColWidthChars = r.u16()
        break
      case REC.COLINFO: {
        const firstCol = r.u16(); const lastCol = r.u16(); const width256 = r.u16(); r.u16() /* xf */
        const opts = r.u16()
        const hidden = (opts & 0x01) !== 0
        for (let c = firstCol; c <= lastCol && c < 16384; c++) {
          if (hidden) hiddenCols[c] = true
          const chars = width256 / 256
          if (chars !== 8.43 && chars > 0) colWidths[c] = colWidthToPx(chars)
        }
        break
      }
      case REC.ROW: {
        const rw = r.u16(); r.u16(); r.u16()
        const heightRaw = r.u16()
        r.u16(); r.u16()
        const flags = r.u16()
        const hidden = (flags & 0x20) !== 0
        if (hidden) hiddenRows[rw] = true
        const heightPt = (heightRaw & 0x7fff) / 20
        if (heightPt > 0) rowHeights[rw] = ptToPx(heightPt)
        break
      }
      case REC.MERGEDCELLS: {
        const n = r.u16()
        for (let k = 0; k < n; k++) {
          const r1 = r.u16(); const r2 = r.u16(); const c1 = r.u16(); const c2 = r.u16()
          merges.push({ r1, c1, r2, c2 })
        }
        break
      }
      case REC.WINDOW2: {
        const opts = r.u16()
        showGridlines = (opts & 0x02) !== 0
        showHeadings = (opts & 0x04) !== 0
        freezePanesFlag = (opts & 0x08) !== 0
        r.u16(); r.u16(); r.u32()
        if (r.left >= 2) { const z = r.u16(); if (z > 0) zoom = z }
        break
      }
      case REC.PANE: {
        const x = r.u16(); const y = r.u16()
        paneSplit = { x, y }
        break
      }
      case REC.BLANK: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        setCell(rw, c, undefined, styleOf(ixfe))
        break
      }
      case REC.MULBLANK: {
        const rw = r.u16(); const firstCol = r.u16()
        const ixfes: number[] = []
        while (r.left > 2) ixfes.push(r.u16())
        r.u16() // lastCol
        for (let k = 0; k < ixfes.length; k++) setCell(rw, firstCol + k, undefined, styleOf(ixfes[k]))
        break
      }
      case REC.NUMBER: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const style = styleOf(ixfe)
        const n = maybeDateAdjust(r.f64(), style)
        setCell(rw, c, String(n), style)
        break
      }
      case REC.RK: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const style = styleOf(ixfe)
        const n = maybeDateAdjust(decodeRK(r.u32()), style)
        setCell(rw, c, String(n), style)
        break
      }
      case REC.MULRK: {
        const rw = r.u16(); const firstCol = r.u16()
        const items: { ixfe: number; rk: number }[] = []
        while (r.left > 2) items.push({ ixfe: r.u16(), rk: r.u32() })
        r.u16()
        for (let k = 0; k < items.length; k++) {
          const style = styleOf(items[k].ixfe)
          setCell(rw, firstCol + k, String(maybeDateAdjust(decodeRK(items[k].rk), style)), style)
        }
        break
      }
      case REC.LABEL: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const text = r.readStr(2, true)
        setCell(rw, c, text, styleOf(ixfe))
        break
      }
      case REC.RSTRING: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const text = r.readStr(2, false)
        // Trailing formatting-run bytes (cRuns * 4) are cosmetic rich-text spans; not modeled.
        setCell(rw, c, text, styleOf(ixfe))
        break
      }
      case REC.LABELSST: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const idx = r.u32()
        setCell(rw, c, sst[idx] ?? '', styleOf(ixfe))
        break
      }
      case REC.BOOLERR: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const val = r.u8(); const isErr = r.u8() === 1
        setCell(rw, c, isErr ? (BIFF_ERR[val] ?? '#N/A') : val ? 'TRUE' : 'FALSE', styleOf(ixfe))
        break
      }
      case REC.SHRFMLA: {
        const r1 = r.u16(); const r2 = r.u16(); const c1 = r.u8(); const c2 = r.u8()
        r.u16() // reserved (cUse)
        const cce = r.u16()
        const rgce = r.bytes(cce)
        const extra = r.left > 0 ? r.rest() : undefined
        groups.push({ r1, c1, r2, c2, rgce, extra, isArray: false })
        break
      }
      case REC.ARRAY: {
        const r1 = r.u16(); const r2 = r.u16(); const c1 = r.u8(); const c2 = r.u8()
        r.u16(); r.u32() // options, reserved
        const cce = r.u16()
        const rgce = r.bytes(cce)
        const extra = r.left > 0 ? r.rest() : undefined
        groups.push({ r1, c1, r2, c2, rgce, extra, isArray: true })
        break
      }
      case REC.FORMULA: {
        const rw = r.u16(); const c = r.u16(); const ixfe = r.u16()
        const style = styleOf(ixfe)
        const cached = readFormulaCachedValue(r)
        const fOpts = r.u16()
        r.u32() // chn, unused
        const cce = r.u16()
        const rgce = r.bytes(cce)
        const extra = r.left > 0 ? r.rest() : undefined
        void fOpts

        let text: string | undefined
        const isPtgExpPointer = rgce.length === 4 && rgce[0] === 0x01
        if (isPtgExpPointer) {
          const group = groups.find((g) => inGroup(g, rw, c))
          if (group) {
            text = decodeFormula(group.rgce, group.extra, ctxFor(rw, c))
            if (text !== undefined && group.isArray) text = '{=' + text + '}'
          }
        } else {
          text = decodeFormula(rgce, extra, ctxFor(rw, c))
          const owningArray = groups.find((g) => g.isArray && inGroup(g, rw, c))
          if (text !== undefined && owningArray) text = '{=' + text + '}'
        }

        let raw: string
        if (text !== undefined) {
          raw = '=' + text
        } else if (cached.kind === 'num') raw = String(maybeDateAdjust(cached.value as number, style))
        else if (cached.kind === 'bool') raw = cached.value ? 'TRUE' : 'FALSE'
        else if (cached.kind === 'error') raw = cached.value as string
        else if (cached.kind === 'string') {
          // Cached value is in the following STRING record; peek ahead.
          const next = records[i + 1]
          if (next && next.type === REC.STRING) {
            const sr = new Rdr(next.data, next.bounds)
            raw = sr.readStr(2, false)
            i++
          } else raw = ''
        } else raw = ''
        setCell(rw, c, raw, style)
        break
      }
      case REC.NOTE: {
        const rw = r.u16(); const c = r.u16(); r.u16() /* flags */; r.u16() /* shapeId */
        const authorLen = r.u16()
        const wide = r.u8() !== 0
        let author = ''
        for (let k = 0; k < authorLen; k++) author += String.fromCharCode(wide ? r.u16() : r.u8())
        const text = pendingNoteTexts.shift() ?? ''
        notes[rw + ',' + c] = { author: author || DEFAULT_AUTHOR, text }
        lastNoteAnchor = { row: rw, col: c }
        void lastNoteAnchor
        break
      }
      case REC.TXO: {
        r.u16() // options
        r.u16(); r.u16(); r.u16(); r.u16() // orientation + reserved
        const textLen = r.u16()
        const fmtLen = r.u16()
        r.u32()
        if (textLen > 0) {
          try { pendingNoteTexts.push(r.readStrKnownLen(textLen)) } catch { pendingNoteTexts.push('') }
        }
        void fmtLen
        break
      }
      default:
        break
    }
  }

  const layout: Partial<SheetLayout> = {}
  layout.colWidths = colWidths
  layout.hiddenCols = hiddenCols
  layout.rowHeights = rowHeights
  layout.hiddenRows = hiddenRows
  layout.merges = merges
  layout.notes = notes
  layout.comments = {}
  layout.showGridlines = showGridlines
  layout.showHeadings = showHeadings
  layout.zoom = zoom
  if (freezePanesFlag && paneSplit) frozen = { rows: paneSplit.y, cols: paneSplit.x }
  if (frozen) layout.frozen = frozen
  void defColWidthChars

  return { cells, layout }
}

// ---------- entry point ----------
export function parseXls(bytes: Uint8Array): WorkbookData {
  let cfb: CFB.CFB$Container
  try {
    cfb = CFB.parse(bytes)
  } catch (e) {
    throw new Error('Not a valid .xls (BIFF8/CFB) file: ' + (e instanceof Error ? e.message : String(e)))
  }
  const entry = cfb.FileIndex.find((e) => /^(workbook|book)$/i.test(e.name))
  if (!entry) throw new Error('Not a valid .xls file: no Workbook stream found.')
  const buf = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content as number[])

  const records = readRecords(buf)
  if (!records.length || records[0].type !== REC.BOF) throw new Error('Not a valid .xls file: missing BOF record.')

  // Globals substream: from the first BOF to its matching EOF.
  let globalsEnd = records.length
  for (let i = 1; i < records.length; i++) {
    if (records[i].type === REC.EOF) { globalsEnd = i; break }
  }
  const globals = parseGlobals(records, globalsEnd)

  const offsetToIndex = new Map<number, number>()
  records.forEach((rec, idx) => offsetToIndex.set(rec.offset, idx))

  const sheetNames = globals.boundSheets.map((b) => b.name)
  const sheets: WorkbookData['sheets'] = []

  globals.boundSheets.forEach((bs) => {
    if (bs.sheetType !== 0) return // chart (2) / macro (1) / VB module (6): not a worksheet - skipped, not rendered as a tab

    const startIdx = offsetToIndex.get(bs.lbPlyPos)
    if (startIdx === undefined || records[startIdx].type !== REC.BOF) {
      sheets.push({ name: bs.name, cells: {}, layout: { ...defaultLayout() } })
      return
    }
    let depth = 0
    let endIdx = records.length
    for (let i = startIdx; i < records.length; i++) {
      if (records[i].type === REC.BOF) depth++
      else if (records[i].type === REC.EOF) { depth--; if (depth === 0) { endIdx = i; break } }
    }

    const { cells, layout } = parseSheet(
      records, startIdx + 1, endIdx, sheetNames, globals.names, globals.externSheets,
      globals.xfStyles, globals.dateMode1904, globals.sst
    )
    sheets.push({ name: bs.name, cells, layout: { ...defaultLayout(), ...layout } })
  })

  if (!sheets.length) sheets.push({ name: 'Sheet1', cells: {}, layout: defaultLayout() })

  return { sheets, activeSheet: 0 }
}

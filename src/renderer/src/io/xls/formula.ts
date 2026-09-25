// Decodes a BIFF8 formula's Ptg RPN token stream (the `rgce` bytes of a FORMULA/ARRAY/NAME/SHRFMLA
// record) into Excel formula text ("=A1+SUM(B1:B2)"). Pure RPN -> infix reconstruction: every
// operator/function just pops its operands off a text stack and pushes the combined text, honoring
// ptgParen tokens for explicit grouping (BIFF never needs *extra* parens beyond what was typed,
// since the token stream already encodes exactly what the user wrote).
//
// Returns `undefined` when a token isn't understood (unmapped function index, array constant with
// an unexpected shape, truncated stream, ...) so the caller can fall back to the formula's cached
// value, exactly like real Excel would show *something* rather than corrupt the formula.
import { colName } from '../../model/types'
import { BIFF_ERR, FUNC_TAB } from './biffConst'

export interface ExternSheetRange {
  first: number // index into ctx.sheetNames, or -1 for "not this workbook" / deleted
  last: number
}

export interface FormulaCtx {
  baseRow: number
  baseCol: number
  sheetNames: string[]
  externSheets: ExternSheetRange[] // indexed by ixti (EXTERNSHEET record)
  names: string[] // 1-based NAME table (names[0] unused)
}

// ---------- little-endian byte reader over a single Uint8Array (the already-flattened rgce +
// any CONTINUE bytes belonging to the same logical record; see biff8.ts's RecordReader). ----------
class ByteReader {
  pos = 0
  constructor(private buf: Uint8Array, public end: number = buf.length) {}
  get left(): number { return this.end - this.pos }
  u8(): number { return this.buf[this.pos++] }
  u16(): number { const v = this.buf[this.pos] | (this.buf[this.pos + 1] << 8); this.pos += 2; return v }
  i16(): number { const v = this.u16(); return v >= 0x8000 ? v - 0x10000 : v }
  u32(): number { const v = (this.buf[this.pos] | (this.buf[this.pos + 1] << 8) | (this.buf[this.pos + 2] << 16) | (this.buf[this.pos + 3] << 24)) >>> 0; this.pos += 4; return v }
  f64(): number {
    const dv = new DataView(this.buf.buffer, this.buf.byteOffset + this.pos, 8)
    this.pos += 8
    return dv.getFloat64(0, true)
  }
  bytes(n: number): Uint8Array { const b = this.buf.subarray(this.pos, this.pos + n); this.pos += n; return b }
  skip(n: number): void { this.pos += n }
}

/** XLUnicodeString-ish read (grbit high-byte flag only; no rich/ext extras) used for ptgStr. */
function readPtgStr(r: ByteReader): string {
  const cch = r.u8()
  const grbit = r.u8()
  const wide = (grbit & 1) !== 0
  let s = ''
  if (wide) {
    for (let i = 0; i < cch; i++) s += String.fromCharCode(r.u16())
  } else {
    for (let i = 0; i < cch; i++) s += String.fromCharCode(r.u8())
  }
  return s
}

function sheetRangeText(ctx: FormulaCtx, first: number, last: number): string {
  const a = ctx.sheetNames[first]
  const b = ctx.sheetNames[last]
  if (a === undefined) return "#REF!"
  const quote = (n: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'")
  if (last === first || b === undefined) return quote(a) + '!'
  return quote(a) + ':' + quote(b) + '!'
}

// Row/col relative-reference decoding (BIFF8 Ref8/RgceArea): the top bits of the column field carry
// the relative flags; when relative, the stored value is a signed delta from the formula's own cell.
function decodeRowCol(rowRaw: number, colRaw: number, baseRow: number, baseCol: number): { row: number; col: number; rowRel: boolean; colRel: boolean } {
  const colRel = (colRaw & 0x8000) !== 0
  const rowRel = (colRaw & 0x4000) !== 0
  let col = colRaw & 0x3fff
  let row = rowRaw
  if (colRel && col > 0x1fff) col -= 0x4000
  if (rowRel && row > 0x7fff) row -= 0x10000
  const MAXR = 1048576
  const MAXC = 16384
  const outRow = rowRel ? (((baseRow + row) % MAXR) + MAXR) % MAXR : row
  const outCol = colRel ? (((baseCol + col) % MAXC) + MAXC) % MAXC : col
  return { row: outRow, col: outCol, rowRel, colRel }
}

function cellText(row: number, col: number, rowRel: boolean, colRel: boolean): string {
  return (colRel ? '' : '$') + colName(col) + (rowRel ? '' : '$') + (row + 1)
}

const BASE_OP: Record<number, string> = {
  0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&',
  0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>',
  0x0f: ' ', 0x10: ',', 0x11: ':'
}

/** Maps a classified ptg id (0x20-0x7F range) down to its reference-class base id (0x20-0x3F). */
function baseOf(ptg: number): number {
  if (ptg < 0x20) return ptg
  return ((ptg - 0x20) & 0x1f) + 0x20
}

interface ArrayConst { rows: number; cols: number; text: string }

export function decodeFormula(rgce: Uint8Array, extra: Uint8Array | undefined, ctx: FormulaCtx): string | undefined {
  try {
    const r = new ByteReader(rgce)
    const stack: string[] = []
    const arraySlots: { placeholderIndex: number }[] = []
    let extraArrays: ArrayConst[] | undefined

    const pop = (): string => {
      const v = stack.pop()
      if (v === undefined) throw new Error('stack underflow')
      return v
    }

    while (r.left > 0) {
      const ptg = r.u8()
      const base = baseOf(ptg)

      switch (base) {
        case 0x01: // ptgExp - array/shared-formula master reference (row,col of the master cell)
        case 0x02: { // ptgTbl - data table
          r.skip(4)
          return undefined // caller should already special-case shared/array formulas before calling us
        }
        case 0x03: case 0x04: case 0x05: case 0x06: case 0x07: case 0x08:
        case 0x09: case 0x0a: case 0x0b: case 0x0c: case 0x0d: case 0x0e:
        case 0x0f: case 0x10: case 0x11: {
          const b = pop(); const a = pop()
          stack.push(a + BASE_OP[base] + b)
          break
        }
        case 0x12: { const a = pop(); stack.push('+' + a); break } // ptgUplus
        case 0x13: { const a = pop(); stack.push('-' + a); break } // ptgUminus
        case 0x14: { const a = pop(); stack.push(a + '%'); break } // ptgPercent
        case 0x15: { const a = pop(); stack.push('(' + a + ')'); break } // ptgParen
        case 0x16: { stack.push(''); break } // ptgMissArg
        case 0x17: { stack.push('"' + readPtgStr(r).replace(/"/g, '""') + '"'); break } // ptgStr
        case 0x18: { r.skip(1); break } // ptgExtend (rare, no textual effect we model)
        case 0x19: { // ptgAttr
          const flags = r.u8()
          if (flags & 0x04) { // AttrChoose: 2-byte count + (count+1) 2-byte offsets follow
            const n = r.u16()
            r.skip((n + 1) * 2)
          } else {
            const data = r.u16()
            if (flags & 0x10) { // AttrSum: wrap the single arg already on the stack as SUM(...)
              const a = pop()
              stack.push('SUM(' + a + ')')
            } else if (flags & 0x40) { // AttrSpace / AttrSpaceVolatile: literal whitespace, cosmetic only
              void data
            }
            // AttrSemi (volatile marker), AttrIf, AttrSkip/Goto: control-flow only, no text effect.
          }
          break
        }
        case 0x1c: { const code = r.u8(); stack.push(BIFF_ERR[code] ?? '#N/A'); break } // ptgErr
        case 0x1d: { const v = r.u8(); stack.push(v ? 'TRUE' : 'FALSE'); break } // ptgBool
        case 0x1e: { stack.push(String(r.u16())); break } // ptgInt
        case 0x1f: { const n = r.f64(); stack.push(formatNumLiteral(n)); break } // ptgNum

        case 0x20: { // ptgArray - the constant's actual data lives in the trailing "extra data" area
          r.skip(7)
          arraySlots.push({ placeholderIndex: stack.length })
          stack.push('\u0000ARRAY\u0000') // resolved once we've read the extra data, below
          break
        }
        case 0x21: { // ptgFunc - fixed-arity built-in
          const iftab = r.u16()
          const meta = FUNC_TAB[iftab]
          if (!meta) return undefined
          const argc = fixedArgCount(iftab)
          const args: string[] = []
          for (let i = 0; i < argc; i++) args.unshift(pop())
          stack.push(meta.name + '(' + args.join(',') + ')')
          break
        }
        case 0x22: { // ptgFuncVar - variable-arity built-in (or add-in via a preceding ptgNameX)
          const cparams = r.u8()
          const iftabRaw = r.u16()
          const iftab = iftabRaw & 0x7fff
          const isCmd = (iftabRaw & 0x8000) !== 0
          const args: string[] = []
          for (let i = 0; i < cparams; i++) args.unshift(pop())
          if (isCmd) {
            // Add-in/command call: the function name was already pushed by a preceding ptgNameX.
            const name = pop()
            stack.push(name + '(' + args.join(',') + ')')
          } else {
            const meta = FUNC_TAB[iftab]
            if (!meta) return undefined
            stack.push(meta.name + '(' + args.join(',') + ')')
          }
          break
        }
        case 0x23: { // ptgName
          const ilbl = r.u16()
          r.skip(2)
          const name = ctx.names[ilbl]
          if (!name) return undefined
          stack.push(name)
          break
        }
        case 0x24: { // ptgRef
          const rowRaw = r.u16(); const colRaw = r.u16()
          const { row, col, rowRel, colRel } = decodeRowCol(rowRaw, colRaw, ctx.baseRow, ctx.baseCol)
          stack.push(cellText(row, col, rowRel, colRel))
          break
        }
        case 0x25: { // ptgArea
          const rwF = r.u16(); const rwL = r.u16(); const colFRaw = r.u16(); const colLRaw = r.u16()
          const a = decodeRowCol(rwF, colFRaw, ctx.baseRow, ctx.baseCol)
          const b = decodeRowCol(rwL, colLRaw, ctx.baseRow, ctx.baseCol)
          stack.push(cellText(a.row, a.col, a.rowRel, a.colRel) + ':' + cellText(b.row, b.col, b.rowRel, b.colRel))
          break
        }
        case 0x26: { r.skip(6); stack.push(''); break } // ptgMemArea
        case 0x27: { r.skip(4); stack.push('#REF!'); break } // ptgMemErr (BIFF8 uses 4 subsequent bytes after class-dependent header; kept minimal)
        case 0x28: { stack.push(''); break } // ptgMemNoMem
        case 0x29: { r.skip(2); break } // ptgMemFunc
        case 0x2a: { r.skip(4); stack.push('#REF!'); break } // ptgRefErr
        case 0x2b: { r.skip(8); stack.push('#REF!'); break } // ptgAreaErr
        case 0x2c: { // ptgRefN (used inside shared formulas / names, base is the formula's own cell)
          const rowRaw = r.u16(); const colRaw = r.u16()
          const { row, col, rowRel, colRel } = decodeRowCol(rowRaw, colRaw, ctx.baseRow, ctx.baseCol)
          stack.push(cellText(row, col, rowRel, colRel))
          break
        }
        case 0x2d: { // ptgAreaN
          const rwF = r.u16(); const rwL = r.u16(); const colFRaw = r.u16(); const colLRaw = r.u16()
          const a = decodeRowCol(rwF, colFRaw, ctx.baseRow, ctx.baseCol)
          const b = decodeRowCol(rwL, colLRaw, ctx.baseRow, ctx.baseCol)
          stack.push(cellText(a.row, a.col, a.rowRel, a.colRel) + ':' + cellText(b.row, b.col, b.rowRel, b.colRel))
          break
        }
        case 0x39: { // ptgNameX - external/add-in name (SUPBOOK + name index)
          r.skip(2) // ixti
          const ilbl = r.u16()
          r.skip(2)
          stack.push(ctx.names[ilbl] ?? ('_xlfn.NAME' + ilbl))
          break
        }
        case 0x3a: { // ptgRef3d
          const ixti = r.u16(); const rowRaw = r.u16(); const colRaw = r.u16()
          const rg = ctx.externSheets[ixti]
          const { row, col, rowRel, colRel } = decodeRowCol(rowRaw, colRaw, ctx.baseRow, ctx.baseCol)
          const prefix = rg ? sheetRangeText(ctx, rg.first, rg.last) : ''
          stack.push(prefix + cellText(row, col, rowRel, colRel))
          break
        }
        case 0x3b: { // ptgArea3d
          const ixti = r.u16()
          const rwF = r.u16(); const rwL = r.u16(); const colFRaw = r.u16(); const colLRaw = r.u16()
          const rg = ctx.externSheets[ixti]
          const a = decodeRowCol(rwF, colFRaw, ctx.baseRow, ctx.baseCol)
          const b = decodeRowCol(rwL, colLRaw, ctx.baseRow, ctx.baseCol)
          const prefix = rg ? sheetRangeText(ctx, rg.first, rg.last) : ''
          stack.push(prefix + cellText(a.row, a.col, a.rowRel, a.colRel) + ':' + cellText(b.row, b.col, b.rowRel, b.colRel))
          break
        }
        case 0x3c: { r.skip(6); stack.push('#REF!'); break } // ptgRefErr3d
        case 0x3d: { r.skip(10); stack.push('#REF!'); break } // ptgAreaErr3d
        default:
          return undefined // unrecognized/unsupported token - bail, caller keeps the cached value
      }
    }

    // Resolve any ptgArray placeholders from the trailing extra-data area (array constants).
    if (arraySlots.length) {
      if (!extra) return undefined
      extraArrays = readArrayConstants(extra, arraySlots.length)
      if (!extraArrays) return undefined
      for (let i = 0; i < arraySlots.length; i++) {
        const idx = stack.indexOf('\u0000ARRAY\u0000')
        if (idx === -1) return undefined
        stack[idx] = extraArrays[i].text
      }
    }

    if (stack.length !== 1) return undefined
    return stack[0]
  } catch {
    return undefined
  }
}

function formatNumLiteral(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n)
  return String(n)
}

/** BIFF8 fixed-arity ptgFunc argument counts for the subset of functions in FUNC_TAB that are ever
 *  encoded via ptgFunc (rather than ptgFuncVar) - i.e. functions with exactly one valid arg count.
 *  Anything not listed here but present in FUNC_TAB is variable-arity (reached via ptgFuncVar only). */
const FIXED_ARGC: Record<number, number> = {
  2: 1, 3: 1, 8: 0, 9: 1, 10: 0, 12: 1, 13: 2, 14: 2, 15: 1, 16: 1, 17: 1, 18: 1, 19: 0, 20: 1,
  21: 1, 22: 1, 23: 1, 24: 1, 25: 1, 26: 1, 27: 2, 30: 2, 31: 3, 32: 1, 33: 1, 34: 0, 35: 0, 38: 1,
  39: 2, 65: 3, 66: 3, 67: 1, 68: 1, 69: 1, 70: 1, 71: 1, 72: 1, 73: 1, 74: 0, 75: 1, 76: 1, 77: 1,
  86: 1, 97: 2, 98: 1, 99: 1, 111: 1, 112: 1, 113: 1, 114: 1, 117: 2, 118: 1, 119: 4, 120: 4,
  121: 1, 126: 1, 127: 1, 128: 1, 129: 1, 130: 1, 131: 1, 140: 1, 141: 1, 142: 3, 143: 4, 144: 4,
  162: 1, 163: 1, 164: 1, 184: 1, 190: 1, 197: 1, 198: 1, 212: 2, 213: 2, 220: 2, 221: 0, 226: 1,
  227: 1, 228: 1, 229: 1, 230: 1, 231: 1, 279: 1, 283: 1, 284: 1, 296: 1, 342: 1, 343: 1
}
function fixedArgCount(iftab: number): number {
  return FIXED_ARGC[iftab] ?? 1
}

/** BIFF8 array-constant "extra data" layout (2.5.198.3 PtgArray extra area): for each ptgArray
 *  token in source order, one entry of {cols-1 (1 byte), rows-1 (2 bytes), cols*rows values}. Each
 *  value is a 1-byte type tag (0=empty,1=number,2=string,4=bool,16=error) followed by its payload. */
function readArrayConstants(extra: Uint8Array, count: number): ArrayConst[] | undefined {
  try {
    const r = new ByteReader(extra)
    const out: ArrayConst[] = []
    for (let k = 0; k < count; k++) {
      const cols = r.u8() + 1
      const rows = r.u16() + 1
      const rowsText: string[] = []
      for (let ri = 0; ri < rows; ri++) {
        const cells: string[] = []
        for (let ci = 0; ci < cols; ci++) {
          const tag = r.u8()
          if (tag === 1) cells.push(formatNumLiteral(r.f64()))
          else if (tag === 2) cells.push('"' + readShortArrayString(r).replace(/"/g, '""') + '"')
          else if (tag === 4) { const v = r.u8(); r.skip(7); cells.push(v ? 'TRUE' : 'FALSE') }
          else if (tag === 16) { const v = r.u8(); r.skip(7); cells.push(BIFF_ERR[v] ?? '#N/A') }
          else { r.skip(8); cells.push('') }
        }
        rowsText.push(cells.join(','))
      }
      out.push({ rows, cols, text: '{' + rowsText.join(';') + '}' })
    }
    return out
  } catch {
    return undefined
  }
}
function readShortArrayString(r: ByteReader): string {
  const cch = r.u16()
  const grbit = r.u8()
  const wide = (grbit & 1) !== 0
  let s = ''
  if (wide) for (let i = 0; i < cch; i++) s += String.fromCharCode(r.u16())
  else for (let i = 0; i < cch; i++) s += String.fromCharCode(r.u8())
  return s
}

// Fast path for writing the cells of big sheets.
//
// exceljs builds an object per cell and per style and then a JSZip archive: about 20 s for 50,000 rows x 21 columns, on the
// UI thread. For sheets over a size threshold the cells (values, formulas and their styles) are written here as XML text
// instead, and spliced into the package exceljs produces for everything else (views, widths, merges, comments, filters...).
import { toFiles, type PkgLike } from './xlsxPostProcess'
import type { CellStyle, BorderSide } from '../model/types'
import { addFuturePrefixes } from '../engine/xlfn'

export interface BulkCellData { raw?: string; value?: unknown; spill?: string; style?: CellStyle }
/** Cells of one sheet grouped by 0-based row, each row sorted by column. */
export type BulkRows = Map<number, { c: number; cell: BulkCellData }[]>

const EXCEL_ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'])

/** Below this many cells a sheet goes through exceljs (fewer moving parts); set CELLBOOK_FAST_WRITE_MIN=0 to force this path. */
export function fastWriteThreshold(): number {
  const env = typeof process !== 'undefined' ? process.env?.CELLBOOK_FAST_WRITE_MIN : undefined
  return env !== undefined && env !== '' ? Number(env) : 15000
}

export function colLetters(c: number): string {
  let s = ''
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}
const colIndexOf = (letters: string): number => [...letters].reduce((v, ch) => v * 26 + ch.charCodeAt(0) - 64, 0) - 1

function esc(s: string): string {
  return s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;')).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
}
const escAttr = (s: string): string => esc(s).replace(/"/g, '&quot;')

const hexToArgb = (hex: string): string => {
  const h = hex.replace('#', '').toUpperCase()
  return 'FF' + (h.length === 6 ? h : h.padStart(6, '0'))
}

// ---------- styles ----------
const BUILTIN_FMT: Record<string, number> = {
  General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '# ?/?': 12, '# ??/??': 13, 'mm-dd-yy': 14, 'd-mmm-yy': 15,
  'd-mmm': 16, 'mmm-yy': 17, 'h:mm AM/PM': 18, 'h:mm:ss AM/PM': 19, 'h:mm': 20, 'h:mm:ss': 21, 'm/d/yy h:mm': 22, '#,##0 ;(#,##0)': 37, '#,##0 ;[Red](#,##0)': 38,
  '#,##0.00;(#,##0.00)': 39, '#,##0.00;[Red](#,##0.00)': 40, 'mm:ss': 45, '[h]:mm:ss': 46, 'mmss.0': 47, '##0.0E+0': 48, '@': 49
}

function borderXml(tag: string, b: BorderSide | undefined): string {
  if (!b || !b.style) return `<${tag}/>`
  return `<${tag} style="${b.style}"><color rgb="${hexToArgb(b.color ?? '#000000')}"/></${tag}>`
}

class StyleAdder {
  private fonts: string[] = []
  private fills: string[] = []
  private borders: string[] = []
  private fmts: string[] = []
  private xfs: string[] = []
  private ids = new Map<string, number>()
  private fontIds = new Map<string, number>()
  private fillIds = new Map<string, number>()
  private borderIds = new Map<string, number>()
  private fmtIds = new Map<string, number>()
  private nextFmt: number
  constructor(private base: { fonts: number; fills: number; borders: number; xfs: number; maxFmt: number }) { this.nextFmt = Math.max(163, base.maxFmt) + 1 }

  private intern(map: Map<string, number>, list: string[], baseCount: number, xml: string): number {
    let id = map.get(xml)
    if (id === undefined) { id = baseCount + list.length; list.push(xml); map.set(xml, id) }
    return id
  }

  /** xf index (into cellXfs) for a style; 0 for "no style". */
  xf(style: CellStyle | undefined): number {
    if (!style) return 0
    const key = JSON.stringify(style)
    const known = this.ids.get(key)
    if (known !== undefined) return known
    let fontId = 0, fillId = 0, borderId = 0, numFmtId = 0
    const font = (style.bold ? '<b/>' : '') + (style.italic ? '<i/>' : '') + (style.strike ? '<strike/>' : '') +
      (style.underline ? (style.underline === 'double' ? '<u val="double"/>' : '<u/>') : '') +
      (style.fontSize ? `<sz val="${style.fontSize}"/>` : '') + (style.color ? `<color rgb="${hexToArgb(style.color)}"/>` : '') +
      (style.fontFamily ? `<name val="${escAttr(style.fontFamily)}"/>` : '')
    if (font) fontId = this.intern(this.fontIds, this.fonts, this.base.fonts, `<font>${font}</font>`)
    if (style.fill) fillId = this.intern(this.fillIds, this.fills, this.base.fills, `<fill><patternFill patternType="solid"><fgColor rgb="${hexToArgb(style.fill)}"/></patternFill></fill>`)
    if (style.borders && (style.borders.top || style.borders.right || style.borders.bottom || style.borders.left)) {
      const b = style.borders
      borderId = this.intern(this.borderIds, this.borders, this.base.borders, `<border>${borderXml('left', b.left)}${borderXml('right', b.right)}${borderXml('top', b.top)}${borderXml('bottom', b.bottom)}<diagonal/></border>`)
    }
    if (style.numFmt && style.numFmt !== 'General') {
      const builtin = BUILTIN_FMT[style.numFmt]
      if (builtin !== undefined) numFmtId = builtin
      else {
        let id = this.fmtIds.get(style.numFmt)
        if (id === undefined) { id = this.nextFmt++; this.fmtIds.set(style.numFmt, id); this.fmts.push(`<numFmt numFmtId="${id}" formatCode="${escAttr(style.numFmt)}"/>`) }
        numFmtId = id
      }
    }
    let al = ''
    if (style.hAlign && style.hAlign !== 'general') al += ` horizontal="${style.hAlign}"`
    if (style.vAlign) al += ` vertical="${style.vAlign === 'middle' ? 'center' : style.vAlign}"`
    if (style.wrap) al += ' wrapText="1"'
    if (style.shrinkToFit) al += ' shrinkToFit="1"'
    if (style.indent) al += ` indent="${style.indent}"`
    if (style.rotation) al += ` textRotation="${style.rotation}"`
    const prot = style.locked === false || style.hidden ? `<protection locked="${style.locked === false ? 0 : 1}" hidden="${style.hidden ? 1 : 0}"/>` : ''
    const attrs = `numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
      (numFmtId ? ' applyNumberFormat="1"' : '') + (fontId ? ' applyFont="1"' : '') + (fillId ? ' applyFill="1"' : '') + (borderId ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1"' : '') + (prot ? ' applyProtection="1"' : '')
    const xml = al || prot ? `<xf ${attrs}>${al ? `<alignment${al}/>` : ''}${prot}</xf>` : `<xf ${attrs}/>`
    const id = this.base.xfs + this.xfs.length
    this.xfs.push(xml)
    this.ids.set(key, id)
    return id
  }

  apply(xml: string): string {
    const addTo = (tag: string, list: string[], src: string): string => {
      if (!list.length) return src
      const re = new RegExp(`<${tag} count="(\\d+)"([^>]*)>`)
      const m = re.exec(src)
      if (!m) return src
      const out = src.replace(re, `<${tag} count="${+m[1] + list.length}"$2>`)
      return out.replace(`</${tag}>`, list.join('') + `</${tag}>`)
    }
    let out = xml
    if (this.fmts.length) {
      if (/<numFmts\b/.test(out)) out = addTo('numFmts', this.fmts, out)
      else out = out.replace(/<fonts\b/, `<numFmts count="${this.fmts.length}">${this.fmts.join('')}</numFmts><fonts`)
    }
    out = addTo('fonts', this.fonts, out)
    out = addTo('fills', this.fills, out)
    out = addTo('borders', this.borders, out)
    out = addTo('cellXfs', this.xfs, out)
    return out
  }
}

function readStyleBase(xml: string): ConstructorParameters<typeof StyleAdder>[0] {
  const count = (tag: string): number => { const m = new RegExp(`<${tag} count="(\\d+)"`).exec(xml); return m ? +m[1] : 0 }
  let maxFmt = 0
  for (const m of xml.matchAll(/<numFmt numFmtId="(\d+)"/g)) maxFmt = Math.max(maxFmt, +m[1])
  return { fonts: count('fonts'), fills: count('fills'), borders: count('borders'), xfs: count('cellXfs'), maxFmt }
}

// ---------- cells ----------
interface StringTable { add(s: string): number; added: string[]; refs: number }

function makeStrings(existingUnique: number): StringTable {
  const ids = new Map<string, number>()
  const added: string[] = []
  return {
    added, refs: 0,
    add(s: string): number {
      this.refs++
      let id = ids.get(s)
      if (id === undefined) {
        id = existingUnique + added.length
        ids.set(s, id)
        added.push(/^\s|\s$|\n|\s\s/.test(s) ? `<si><t xml:space="preserve">${esc(s)}</t></si>` : `<si><t>${esc(s)}</t></si>`)
      }
      return id
    }
  }
}

function cellXml(ref: string, cell: BulkCellData, xf: number, strings: StringTable): string {
  const s = xf ? ` s="${xf}"` : ''
  const raw = cell.raw
  if (raw === undefined || raw === '') return xf ? `<c r="${ref}"${s}/>` : ''
  const c0 = raw.charCodeAt(0)
  if (c0 === 39) return `<c r="${ref}"${s} t="s"><v>${strings.add(raw.slice(1))}</v></c>` // 'quoted text
  if (c0 === 61) {
    const f = esc(addFuturePrefixes(raw.slice(1)))
    const v = cell.value
    const err = v && typeof v === 'object' ? (v as { error: string }).error : undefined
    const fx = cell.spill ? `<f t="array" ref="${cell.spill}">${f}</f>` : `<f>${f}</f>`
    if (v === null || v === undefined || (err !== undefined && !EXCEL_ERRORS.has(err))) return `<c r="${ref}"${s}>${fx}</c>`
    if (err !== undefined) return `<c r="${ref}"${s} t="e">${fx}<v>${esc(err)}</v></c>`
    if (typeof v === 'number') return `<c r="${ref}"${s}>${fx}<v>${v}</v></c>`
    if (typeof v === 'boolean') return `<c r="${ref}"${s} t="b">${fx}<v>${v ? 1 : 0}</v></c>`
    return `<c r="${ref}"${s} t="str">${fx}<v>${esc(String(v))}</v></c>`
  }
  if (raw === 'TRUE' || raw === 'FALSE') return `<c r="${ref}"${s} t="b"><v>${raw === 'TRUE' ? 1 : 0}</v></c>`
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw)) return `<c r="${ref}"${s}><v>${Number(raw)}</v></c>`
  return `<c r="${ref}"${s} t="s"><v>${strings.add(raw)}</v></c>`
}

/** Splices the bulk rows of each sheet into the package. `sheets[i]` belongs to xl/worksheets/sheet{i+1}.xml. */
export function injectBulkCells(zip: PkgLike, sheets: (BulkRows | undefined)[]): PkgLike {
  if (!sheets.some(Boolean)) return zip
  const files = toFiles(zip)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  const stylesXml = dec.decode(files['xl/styles.xml'])
  const styles = new StyleAdder(readStyleBase(stylesXml))
  const sstXml = files['xl/sharedStrings.xml'] ? dec.decode(files['xl/sharedStrings.xml']) : null
  const unique = sstXml ? +(/uniqueCount="(\d+)"/.exec(sstXml)?.[1] ?? 0) : 0
  const strings = makeStrings(unique)

  sheets.forEach((rows, i) => {
    if (!rows) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    let xml = dec.decode(files[path])
    // What exceljs kept: rows with heights/hidden flags and the few special cells (notes, merges).
    const existing = new Map<number, { attrs: string; cells: { c: number; xml: string }[] }>()
    const sd = /<sheetData(?:\/>|>([\s\S]*?)<\/sheetData>)/.exec(xml)
    if (!sd) return
    const body = sd[1] ?? ''
    for (const m of body.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const r = +(/\br="(\d+)"/.exec(m[1])?.[1] ?? 0) - 1
      const attrs = m[1].replace(/\s*\/$/, '').replace(/\s+spans="[^"]*"/, '').replace(/\s+r="\d+"/, '')
      const cells: { c: number; xml: string }[] = []
      for (const cm of (m[2] ?? '').matchAll(/<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) cells.push({ c: colIndexOf(cm[1]), xml: cm[0] })
      existing.set(r, { attrs, cells })
    }
    const allRows = new Set<number>([...existing.keys(), ...rows.keys()])
    const order = [...allRows].sort((a, b) => a - b)
    let minR = Infinity, maxR = -1, minC = Infinity, maxC = -1
    const out: string[] = []
    for (const r of order) {
      const ex = existing.get(r)
      const bulk = rows.get(r) ?? []
      const parts: { c: number; xml: string }[] = []
      for (const { c, cell } of bulk) {
        const x = cellXml(colLetters(c) + (r + 1), cell, styles.xf(cell.style), strings)
        if (x) parts.push({ c, xml: x })
      }
      if (ex) for (const e of ex.cells) if (!parts.some((p) => p.c === e.c)) parts.push(e)
      parts.sort((a, b) => a.c - b.c)
      if (!parts.length && !ex?.attrs.trim()) continue
      if (parts.length) { minR = Math.min(minR, r); maxR = Math.max(maxR, r); minC = Math.min(minC, parts[0].c); maxC = Math.max(maxC, parts[parts.length - 1].c) }
      const attrs = ex?.attrs ?? ''
      out.push(parts.length ? `<row r="${r + 1}"${attrs}>${parts.map((p) => p.xml).join('')}</row>` : `<row r="${r + 1}"${attrs}/>`)
    }
    xml = xml.slice(0, sd.index) + `<sheetData>${out.join('')}</sheetData>` + xml.slice(sd.index + sd[0].length)
    if (maxR >= 0) {
      const dim = `${colLetters(minC)}${minR + 1}:${colLetters(maxC)}${maxR + 1}`
      xml = /<dimension\b/.test(xml) ? xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${dim}"/>`) : xml.replace(/(<sheetViews|<sheetFormatPr|<sheetData)/, `<dimension ref="${dim}"/>$1`)
    }
    files[path] = enc.encode(xml)
  })

  files['xl/styles.xml'] = enc.encode(styles.apply(stylesXml))
  if (strings.added.length) {
    if (sstXml) {
      const count = +(/\bcount="(\d+)"/.exec(sstXml)?.[1] ?? 0)
      files['xl/sharedStrings.xml'] = enc.encode(
        sstXml.replace(/\bcount="\d+"/, `count="${count + strings.refs}"`).replace(/uniqueCount="\d+"/, `uniqueCount="${unique + strings.added.length}"`).replace('</sst>', strings.added.join('') + '</sst>')
      )
    } else {
      // exceljs writes no string table when the workbook has no strings of its own
      files['xl/sharedStrings.xml'] = enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.refs}" uniqueCount="${strings.added.length}">${strings.added.join('')}</sst>`)
      const ct = dec.decode(files['[Content_Types].xml'])
      files['[Content_Types].xml'] = enc.encode(ct.replace('</Types>', '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'))
      const rels = dec.decode(files['xl/_rels/workbook.xml.rels'])
      files['xl/_rels/workbook.xml.rels'] = enc.encode(rels.replace('</Relationships>', '<Relationship Id="rIdBulkSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'))
    }
  }
  return { files }
}

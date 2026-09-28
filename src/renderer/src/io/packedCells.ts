// Compact form of a big sheet's cells for handing the parse worker's result to the UI thread.
// Structured-cloning a million `{ raw, style }` objects took about as long as parsing the file; typed arrays are
// transferred without copying, and the raw strings travel as one big string.
import type { CellStyle } from '../model/types'

export interface PackedCells {
  r: Int32Array
  c: Int32Array
  /** Index into `styles`, or -1 for an unstyled cell. */
  sid: Int32Array
  /** Length of the cell's raw text inside `blob`, or -1 when the cell has none. */
  len: Int32Array
  blob: string
  styles: CellStyle[]
}

interface CellLike { raw?: string; style?: CellStyle }

export function packCells(cells: Record<string, CellLike>): PackedCells {
  const keys = Object.keys(cells)
  const n = keys.length
  const r = new Int32Array(n), c = new Int32Array(n), sid = new Int32Array(n), len = new Int32Array(n)
  const styles: CellStyle[] = []
  const ids = new Map<CellStyle, number>()
  const parts: string[] = []
  for (let i = 0; i < n; i++) {
    const k = keys[i]
    const comma = k.indexOf(',')
    r[i] = +k.slice(0, comma)
    c[i] = +k.slice(comma + 1)
    const cell = cells[k]
    if (cell.raw) { len[i] = cell.raw.length; parts.push(cell.raw) } else len[i] = -1
    const st = cell.style
    if (st && Object.keys(st).length) {
      let id = ids.get(st)
      if (id === undefined) { id = styles.length; styles.push(st); ids.set(st, id) }
      sid[i] = id
    } else sid[i] = -1
  }
  return { r, c, sid, len, blob: parts.join(''), styles }
}

/** Calls `fn` for every cell of a packed sheet, in stored order. */
export function eachPacked(p: PackedCells, fn: (r: number, c: number, raw: string | undefined, style: CellStyle | undefined) => void): void {
  let off = 0
  for (let i = 0; i < p.r.length; i++) {
    const L = p.len[i]
    let raw: string | undefined
    if (L >= 0) { raw = p.blob.substr(off, L); off += L }
    fn(p.r[i], p.c[i], raw, p.sid[i] >= 0 ? p.styles[p.sid[i]] : undefined)
  }
}

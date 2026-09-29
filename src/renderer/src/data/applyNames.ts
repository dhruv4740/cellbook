// Formulas > Define Name > Apply Names: replaces cell / range references in formulas with the defined names that point at them.
import { store, SheetState } from '../model/store'

const REF = /(?<![\w.$!'])((?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(!])/g

const norm = (ref: string, home: string): string => {
  const bang = ref.lastIndexOf('!')
  const sheet = bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : home
  return `${sheet.toLowerCase()}!${ref.slice(bang + 1).replace(/\$/g, '').toUpperCase()}`
}

function masked(f: string): string {
  let out = '', inStr = false
  for (let i = 0; i < f.length; i++) {
    if (f[i] === '"') { if (inStr && f[i + 1] === '"') { out += '  '; i++; continue } inStr = !inStr; out += ' '; continue }
    out += inStr ? ' ' : f[i]
  }
  return out
}

/** Which names can be applied: those that point at exactly one cell or one range. */
export function applicableNames(): { name: string; ref: string }[] {
  return store.names.filter((n) => !n.hidden && /^(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/.test(n.ref)).map((n) => ({ name: n.name, ref: n.ref }))
}

export function applyNames(sheet: SheetState, names: string[]): number {
  const wanted = new Map<string, string>()
  for (const n of applicableNames()) if (names.includes(n.name)) wanted.set(norm(n.ref, sheet.name), n.name)
  const used = store.usedRange(sheet)
  if (!wanted.size || !used) return 0
  const edits: { r: number; c: number; raw: string }[] = []
  for (let r = used.r1; r <= used.r2; r++) for (let c = used.c1; c <= used.c2; c++) {
    const raw = store.getRaw(sheet, r, c)
    if (!raw.startsWith('=')) continue
    const m = masked(raw)
    let out = '', last = 0
    REF.lastIndex = 0
    for (let x = REF.exec(m); x; x = REF.exec(m)) {
      const name = wanted.get(norm(raw.slice(x.index, x.index + x[0].length), sheet.name))
      if (!name) continue
      out += raw.slice(last, x.index) + name
      last = x.index + x[0].length
    }
    if (last) edits.push({ r, c, raw: out + raw.slice(last) })
  }
  if (edits.length) store.setCells(sheet, edits, 'Apply Names')
  return edits.length
}

// Review > Spelling (F7): an offline English dictionary (nspell + Hunspell en-US, loaded on first use).
import type NSpell from 'nspell'
import { store, SheetState } from '../model/store'
import { normRange } from '../model/types'

let spellP: Promise<NSpell> | null = null
const CUSTOM_KEY = 'cellbook.spell.words'
const ignoredAll = new Set<string>()

export function loadSpell(): Promise<NSpell> {
  spellP ??= (async () => {
    const [{ default: nspell }, aff, dic] = await Promise.all([
      import('nspell'),
      import('../../../../node_modules/dictionary-en/index.aff?raw').then((m) => m.default),
      import('../../../../node_modules/dictionary-en/index.dic?raw').then((m) => m.default)
    ])
    const sp = nspell({ aff, dic })
    try { for (const w of JSON.parse(localStorage.getItem(CUSTOM_KEY) ?? '[]') as string[]) sp.add(w) } catch { /* no custom words yet */ }
    return sp
  })()
  return spellP
}

export async function addToDictionary(word: string): Promise<void> {
  const sp = await loadSpell()
  sp.add(word)
  try {
    const words = new Set<string>(JSON.parse(localStorage.getItem(CUSTOM_KEY) ?? '[]'))
    words.add(word)
    localStorage.setItem(CUSTOM_KEY, JSON.stringify([...words]))
  } catch { /* storage unavailable: the word still counts for this session */ }
}

export const ignoreAll = (word: string): void => { ignoredAll.add(word.toLowerCase()) }
export const resetIgnored = (): void => ignoredAll.clear()

export interface Misspelling { r: number; c: number; word: string; index: number; text: string }

const WORD = /[A-Za-z][A-Za-z'’]*/g

/** Words that Excel's default options never flag: ALL CAPS, anything glued to digits, URLs and e-mail addresses. */
function skippable(text: string, m: RegExpMatchArray): boolean {
  const w = m[0]
  const i = m.index ?? 0
  if (w.length < 2 || w === w.toUpperCase()) return true
  if (/[0-9_@]/.test(text[i - 1] ?? '') || /[0-9_@]/.test(text[i + w.length] ?? '')) return true
  return false
}

/** The next misspelled word at or after (r, c, index) inside `bounds`, scanning row by row. */
export function findMisspelling(sp: NSpell, sheet: SheetState, bounds: { r1: number; c1: number; r2: number; c2: number }, from: { r: number; c: number; index: number }): Misspelling | null {
  const b = normRange(bounds)
  for (let r = Math.max(b.r1, from.r); r <= b.r2; r++) {
    for (let c = r === from.r ? Math.max(b.c1, from.c) : b.c1; c <= b.c2; c++) {
      if (typeof store.getValue(sheet, r, c) !== 'string') continue
      const raw = store.getRaw(sheet, r, c)
      if (!raw || raw.startsWith('=')) continue
      const text = raw.startsWith("'") ? raw.slice(1) : raw
      const start = r === from.r && c === from.c ? from.index : 0
      const scan = text.replace(/(?:https?:\/\/|www\.)\S+|\S+@\S+/gi, (u) => ' '.repeat(u.length)) // URLs and e-mail addresses
      for (const m of scan.matchAll(WORD)) {
        if ((m.index ?? 0) < start || skippable(scan, m)) continue
        const w = m[0].replace(/’/g, "'")
        if (ignoredAll.has(w.toLowerCase()) || sp.correct(w)) continue
        return { r, c, word: m[0], index: m.index ?? 0, text }
      }
    }
  }
  return null
}

export function suggestions(sp: NSpell, word: string): string[] {
  const w0 = word.replace(/’/g, "'")
  // nspell doesn't try swapped neighbours ("teh" -> "the"), the most common typing slip: offer those first.
  const swaps: string[] = []
  for (let i = 0; i < w0.length - 1; i++) {
    const t = w0.slice(0, i) + w0[i + 1] + w0[i] + w0.slice(i + 2)
    if (t !== w0 && sp.correct(t.toLowerCase())) swaps.push(t.toLowerCase())
  }
  const s = [...new Set([...swaps, ...sp.suggest(w0)])].slice(0, 8)
  // keep the case of the original ("teh" -> "the", "Teh" -> "The")
  return s.map((x) => (word[0] === word[0].toUpperCase() && word.slice(1) === word.slice(1).toLowerCase() ? x[0].toUpperCase() + x.slice(1) : x))
}

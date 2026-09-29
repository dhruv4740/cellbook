// Review > Spelling (F7): steps through the misspelled words in the selection, or from the active cell to the end of the sheet.
import { useEffect, useRef, useState } from 'react'
import type NSpell from 'nspell'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { MAX_COLS, MAX_ROWS, normRange, colName } from '../model/types'
import { addToDictionary, findMisspelling, ignoreAll, loadSpell, Misspelling, resetIgnored, suggestions } from '../data/spell'

export function SpellingDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const sel = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
  const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2
  const used = store.usedRange(sheet)
  const bounds = useRef(multi ? sel : { r1: 0, c1: 0, r2: Math.min(used?.r2 ?? 0, MAX_ROWS - 1), c2: Math.min(used?.c2 ?? 0, MAX_COLS - 1) })
  const spRef = useRef<NSpell | null>(null)
  const changeAllRef = useRef(new Map<string, string>())
  const [miss, setMiss] = useState<Misspelling | null>(null)
  const [done, setDone] = useState(false)
  const [sugg, setSugg] = useState<string[]>([])
  const [changeTo, setChangeTo] = useState('')

  const advance = (from: { r: number; c: number; index: number }): void => {
    const sp = spRef.current!
    let pos = from
    for (let guard = 0; guard < 100000; guard++) {
      const m = findMisspelling(sp, sheet, bounds.current, pos)
      if (!m) { setMiss(null); setDone(true); return }
      const auto = changeAllRef.current.get(m.word.toLowerCase())
      if (auto === undefined) {
        const s = suggestions(sp, m.word)
        setMiss(m); setSugg(s); setChangeTo(s[0] ?? '')
        store.setSelection({ ranges: [{ r1: m.r, c1: m.c, r2: m.r, c2: m.c }], active: { r: m.r, c: m.c } })
        return
      }
      pos = { r: m.r, c: m.c, index: replaceWord(m, auto) }
    }
  }

  const replaceWord = (m: Misspelling, to: string): number => {
    const raw = store.getRaw(sheet, m.r, m.c)
    const prefix = raw.startsWith("'") ? "'" : ''
    store.setRaw(sheet, m.r, m.c, prefix + m.text.slice(0, m.index) + to + m.text.slice(m.index + m.word.length))
    return m.index + to.length
  }

  useEffect(() => {
    resetIgnored()
    let live = true
    loadSpell().then((sp) => {
      if (!live) return
      spRef.current = sp
      advance(multi ? { r: sel.r1, c: sel.c1, index: 0 } : { r: sheet.selection.active.r, c: sheet.selection.active.c, index: 0 })
    })
    return () => { live = false }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const after = (m: Misspelling, index: number): void => advance({ r: m.r, c: m.c, index })
  const btn = (label: string, onClick: () => void, disabled = !miss): React.ReactElement => (
    <button className="xl-btn" style={{ display: 'block', width: '100%', marginBottom: 6 }} disabled={disabled} onClick={onClick}>{label}</button>
  )

  if (done) {
    return (
      <Dialog title="Spelling" width={340} onCancel={onClose} onOk={onClose}>
        <div style={{ fontSize: 12 }}>The spelling check is complete. You&rsquo;re good to go!</div>
      </Dialog>
    )
  }
  return (
    <Dialog title="Spelling: English (United States)" width={560} onCancel={onClose} cancelLabel="Close">
      <div style={{ display: 'flex', gap: 12, fontSize: 12 }}>
        <div style={{ flex: 1 }}>
          <div>Not in Dictionary:</div>
          <div style={{ border: '1px solid #c6c6c6', minHeight: 54, padding: 4, background: '#fff', whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>
            {miss ? (
              <>
                {miss.text.slice(0, miss.index)}<span style={{ color: '#d00000', fontWeight: 600 }}>{miss.word}</span>{miss.text.slice(miss.index + miss.word.length)}
              </>
            ) : 'Checking…'}
          </div>
          <div style={{ color: '#666', margin: '2px 0 6px' }}>{miss ? `Cell ${colName(miss.c)}${miss.r + 1}` : ''}</div>
          <div>Suggestions:</div>
          <select size={7} style={{ width: '100%' }} value={sugg.includes(changeTo) ? changeTo : ''} onChange={(e) => setChangeTo(e.target.value)} onDoubleClick={() => miss && after(miss, replaceWord(miss, changeTo))}>
            {sugg.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 70 }}>Change to:</label><input value={changeTo} onChange={(e) => setChangeTo(e.target.value)} style={{ flex: 1 }} /></div>
        </div>
        <div style={{ width: 150, paddingTop: 16 }}>
          {btn('Ignore Once', () => miss && after(miss, miss.index + miss.word.length))}
          {btn('Ignore All', () => { if (!miss) return; ignoreAll(miss.word); after(miss, miss.index + miss.word.length) })}
          {btn('Add to Dictionary', () => { if (!miss) return; void addToDictionary(miss.word).then(() => after(miss, miss.index + miss.word.length)) })}
          <div style={{ height: 8 }} />
          {btn('Change', () => miss && after(miss, replaceWord(miss, changeTo)), !miss || !changeTo)}
          {btn('Change All', () => { if (!miss) return; changeAllRef.current.set(miss.word.toLowerCase(), changeTo); after(miss, replaceWord(miss, changeTo)) }, !miss || !changeTo)}
        </div>
      </div>
    </Dialog>
  )
}

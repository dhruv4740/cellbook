// Insert > Symbol: pick a character and put it into the cell being edited (or start editing the active cell).
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { getUi, setUi } from '../chrome/ui'

const SUBSETS: Record<string, [number, number][]> = {
  'Basic Latin': [[0x21, 0x7e]],
  'Latin-1 Supplement': [[0xa1, 0xff]],
  'Latin Extended-A': [[0x100, 0x17f]],
  'Greek and Coptic': [[0x391, 0x3a9], [0x3b1, 0x3c9]],
  'Cyrillic': [[0x410, 0x44f]],
  'General Punctuation': [[0x2010, 0x2027], [0x2030, 0x205e]],
  'Currency Symbols': [[0x24, 0x24], [0xa2, 0xa5], [0x20a0, 0x20bf]],
  'Letterlike Symbols': [[0x2100, 0x214f]],
  'Number Forms': [[0x2150, 0x218b]],
  'Arrows': [[0x2190, 0x21ff]],
  'Mathematical Operators': [[0x2200, 0x22ff]],
  'Box Drawing': [[0x2500, 0x257f]],
  'Geometric Shapes': [[0x25a0, 0x25ff]],
  'Miscellaneous Symbols': [[0x2600, 0x26ff]],
  'Dingbats': [[0x2701, 0x27bf]]
}
const chars = (subset: string): string[] => SUBSETS[subset].flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => String.fromCodePoint(a + i)))

export function SymbolDialog({ onClose }: { onClose: () => void }) {
  const [subset, setSubset] = useState('Basic Latin')
  const [pick, setPick] = useState<string | null>(null)
  const [recent, setRecent] = useState<string[]>(() => { try { return JSON.parse(localStorage.getItem('cellbook.symbols.recent') ?? '[]') } catch { return [] } })

  const insert = (ch: string): void => {
    const ui = getUi()
    const at = store.active.selection.active
    if (ui.edit) {
      const caret = ui.edit.caret ?? ui.edit.text.length
      setUi({ edit: { ...ui.edit, text: ui.edit.text.slice(0, caret) + ch + ui.edit.text.slice(caret), caret: caret + ch.length } })
    } else {
      const raw = store.getRaw(store.active, at.r, at.c)
      const text = raw.startsWith('=') ? raw : raw
      setUi({ edit: { r: at.r, c: at.c, text: text + ch, caret: text.length + ch.length, source: 'cell', mode: 'Edit' } })
    }
    const next = [ch, ...recent.filter((x) => x !== ch)].slice(0, 16)
    setRecent(next)
    try { localStorage.setItem('cellbook.symbols.recent', JSON.stringify(next)) } catch { /* storage unavailable */ }
  }
  const cell = (ch: string, key: string): React.ReactElement => (
    <div key={key} onClick={() => setPick(ch)} onDoubleClick={() => { insert(ch); onClose() }} title={'U+' + ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}
      style={{ width: 26, height: 26, display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid ' + (pick === ch ? '#217346' : '#e4e4e4'), background: pick === ch ? '#e2f0e8' : '#fff', cursor: 'default', fontSize: 15 }}>{ch}</div>
  )
  return (
    <Dialog title="Symbol" width={520} onCancel={onClose} cancelLabel="Close" onOk={() => { if (pick) { insert(pick); onClose() } }} okLabel="Insert">
      <div style={{ fontSize: 12 }}>
        <div className="xl-field-row" style={{ marginBottom: 6 }}>
          <label style={{ width: 60 }}>Subset:</label>
          <select value={subset} onChange={(e) => { setSubset(e.target.value); setPick(null) }}>{Object.keys(SUBSETS).map((s) => <option key={s}>{s}</option>)}</select>
        </div>
        <div style={{ display: 'flex', flexWrap: 'wrap', width: 27 * 16, height: 190, overflow: 'auto', alignContent: 'flex-start' }}>
          {chars(subset).map((ch, i) => cell(ch, String(i)))}
        </div>
        <div style={{ marginTop: 8 }}>Recently used symbols:</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', minHeight: 28 }}>{recent.map((ch, i) => cell(ch, 'r' + i))}</div>
        <div style={{ marginTop: 6, color: '#666' }}>{pick ? `Character code: ${pick.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')} (hex)` : ''}</div>
      </div>
    </Dialog>
  )
}

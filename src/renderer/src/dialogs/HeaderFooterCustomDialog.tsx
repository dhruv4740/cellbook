// Page Setup > Header/Footer > Custom Header.../Custom Footer...: the 3-section (Left/Center/Right)
// editor with the insert-code toolbar (&P page, &N pages, &D date, &T time, &F file, &A sheet).
import { useRef, useState } from 'react'
import { Dialog } from '../components/Dialog'
import { HeaderFooterText } from '../model/types'
import './PageSetupDialog.css'

const CODES: { label: string; code: string; title: string }[] = [
  { label: 'A', code: '&A', title: 'Sheet Name' },
  { label: '#', code: '&P', title: 'Page Number' },
  { label: '##', code: '&N', title: 'Number of Pages' },
  { label: '📅', code: '&D', title: 'Date' },
  { label: '🕐', code: '&T', title: 'Time' },
  { label: '📁', code: '&Z&F', title: 'File Path and Name' },
  { label: '📄', code: '&F', title: 'File Name' },
  { label: 'B', code: '&B', title: 'Bold' },
  { label: 'I', code: '&I', title: 'Italic' }
]

export function HeaderFooterCustomDialog({
  title, initial, onOk, onCancel
}: { title: string; initial: HeaderFooterText; onOk: (v: HeaderFooterText) => void; onCancel: () => void }) {
  const [v, setV] = useState<HeaderFooterText>(initial)
  const [focused, setFocused] = useState<'left' | 'center' | 'right'>('left')
  const refs = { left: useRef<HTMLTextAreaElement>(null), center: useRef<HTMLTextAreaElement>(null), right: useRef<HTMLTextAreaElement>(null) }

  const insert = (code: string) => {
    const key = focused
    const el = refs[key].current
    const cur = v[key]
    const pos = el?.selectionStart ?? cur.length
    const next = cur.slice(0, pos) + code + cur.slice(pos)
    setV({ ...v, [key]: next })
    requestAnimationFrame(() => { el?.focus(); el?.setSelectionRange(pos + code.length, pos + code.length) })
  }

  return (
    <Dialog title={title} width={480} onOk={() => onOk(v)} onCancel={onCancel}>
      <div className="xl-ps-hfcustom">
        <div className="xl-ps-hfcustom-toolbar">
          {CODES.map((c) => (
            <button key={c.code} className="xl-ps-hfcode-btn" title={c.title} onClick={() => insert(c.code)}>{c.label}</button>
          ))}
        </div>
        <div className="xl-ps-hfcustom-cols">
          {(['left', 'center', 'right'] as const).map((key) => (
            <div className="xl-ps-hfcustom-col" key={key}>
              <div className="xl-ps-label">{key === 'left' ? 'Left section:' : key === 'center' ? 'Center section:' : 'Right section:'}</div>
              <textarea
                ref={refs[key]}
                className="xl-ps-hf-textarea"
                value={v[key]}
                onFocus={() => setFocused(key)}
                onChange={(e) => setV({ ...v, [key]: e.target.value })}
              />
            </div>
          ))}
        </div>
      </div>
    </Dialog>
  )
}

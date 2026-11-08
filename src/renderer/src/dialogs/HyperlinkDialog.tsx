// Insert > Link (Ctrl+K): link to a web page / file / email, or to a place in this workbook.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { linkAt, setLink } from '../grid/links'

export function HyperlinkDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const { r, c } = sheet.selection.active
  const existing = linkAt(sheet, r, c)
  const [kind, setKind] = useState<'web' | 'place' | 'email'>(existing?.location ? 'place' : existing?.target?.startsWith('mailto:') ? 'email' : 'web')
  const [text, setText] = useState(store.getDisplay(sheet, r, c).text)
  const [addr, setAddr] = useState(existing?.target?.replace(/^mailto:/, '') ?? '')
  const loc0 = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(existing?.location ?? '')
  const [placeSheet, setPlaceSheet] = useState(loc0 ? (loc0[1] ?? loc0[2]).replace(/''/g, "'") : sheet.name)
  const [placeRef, setPlaceRef] = useState(loc0 ? loc0[3] : 'A1')
  const [tip, setTip] = useState(existing?.tooltip ?? '')
  const [err, setErr] = useState<string | null>(null)
  const ok = () => {
    let link
    if (kind === 'place') {
      if (!/^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(placeRef.trim())) { setErr('Type a cell reference, for example A1.'); return }
      const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(placeSheet) ? placeSheet : `'${placeSheet.replace(/'/g, "''")}'`
      link = { location: `${q}!${placeRef.trim().toUpperCase()}` }
    } else {
      const a = addr.trim()
      if (!a) { setErr(kind === 'email' ? 'Type an e-mail address.' : 'Type an address.'); return }
      const target = kind === 'email' ? 'mailto:' + a.replace(/^mailto:/i, '') : /^[a-z]+:\/\//i.test(a) || /^[A-Za-z]:[\\/]|^\\\\/.test(a) ? a : /^www\./i.test(a) || /\.[a-z]{2,}(\/|$)/i.test(a) ? 'https://' + a : a
      link = { target }
    }
    const withTip = tip.trim() ? { ...link, tooltip: tip.trim() } : link
    const shown = text.trim() || (kind === 'place' ? link.location! : kind === 'email' ? addr.trim() : addr.trim())
    setLink(sheet, r, c, withTip, shown)
    onClose()
  }
  const tabBtn = (k: typeof kind, label: string) => (
    <button className={'xl-btn' + (kind === k ? ' primary' : '')} onClick={() => setKind(k)} style={{ display: 'block', width: '100%', marginBottom: 4 }}>{label}</button>
  )
  return (
    <Dialog title={existing ? 'Edit Hyperlink' : 'Insert Hyperlink'} width={520} onCancel={onClose} onOk={ok}
      footerLeft={existing ? <button className="xl-btn" onClick={() => { setLink(sheet, r, c, null); onClose() }}>Remove Link</button> : undefined}>
      <div style={{ display: 'flex', gap: 12, fontSize: 12 }}>
        <div style={{ width: 130 }}>
          <div style={{ marginBottom: 4 }}>Link to:</div>
          {tabBtn('web', 'Existing File or Web Page')}
          {tabBtn('place', 'Place in This Document')}
          {tabBtn('email', 'E-mail Address')}
        </div>
        <div style={{ flex: 1 }}>
          <div className="xl-field-row"><label style={{ width: 110 }}>Text to display:</label><input value={text} onChange={(e) => setText(e.target.value)} style={{ flex: 1 }} /></div>
          {kind !== 'place' && (
            <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 110 }}>{kind === 'email' ? 'E-mail address:' : 'Address:'}</label><input autoFocus value={addr} onChange={(e) => setAddr(e.target.value)} style={{ flex: 1 }} placeholder={kind === 'email' ? 'name@example.com' : 'https://… or C:\\folder\\file.xlsx'} /></div>
          )}
          {kind === 'place' && (
            <>
              <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 110 }}>Cell reference:</label><input autoFocus value={placeRef} onChange={(e) => setPlaceRef(e.target.value)} style={{ flex: 1 }} /></div>
              <div style={{ marginTop: 6 }}>Or select a place in this document:</div>
              <select size={6} style={{ width: '100%' }} value={placeSheet} onChange={(e) => setPlaceSheet(e.target.value)}>
                {store.sheets.map((s) => <option key={s.uid} value={s.name}>{s.name}</option>)}
              </select>
            </>
          )}
          <div className="xl-field-row" style={{ marginTop: 6 }}><label style={{ width: 110 }}>ScreenTip:</label><input value={tip} onChange={(e) => setTip(e.target.value)} style={{ flex: 1 }} /></div>
          {err && <div style={{ color: '#a4262c', marginTop: 6 }}>{err}</div>}
        </div>
      </div>
    </Dialog>
  )
}

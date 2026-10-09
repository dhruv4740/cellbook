// Data > From Web: enter a URL, pick a table from the page, load it into a new sheet.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { useStore } from '../model/store'
import { parseWebTables, type WebTable } from '../io/webData'
import { showToast } from '../chrome/ui'

export function WebDialog({ onClose }: { onClose: () => void }) {
  const store = useStore()
  const [url, setUrl] = useState('https://')
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState('')
  const [tables, setTables] = useState<WebTable[] | null>(null)
  const [pick, setPick] = useState(0)

  const fetchIt = async (): Promise<void> => {
    setBusy(true); setErr(''); setTables(null)
    try {
      const r = await window.api?.webFetch(url.trim())
      if (!r) throw new Error('Web access is not available.')
      const t = parseWebTables(r.body, r.contentType)
      if (!t.length) throw new Error('No tables were found at that address.')
      setTables(t); setPick(0)
    } catch (e) { setErr((e as Error).message.replace(/^Error invoking remote method '[^']+': (Error: )?/, '')) }
    setBusy(false)
  }
  const load = (): void => {
    const t = tables?.[pick]
    if (!t) return
    store.addSheet(t.name.slice(0, 31).replace(/[\/?*[\]:]/g, ' '))
    const cells = t.rows.flatMap((row, r) => row.map((raw, c) => ({ r, c, raw })).filter((x) => x.raw !== ''))
    store.setCells(store.active, cells, 'From Web')
    onClose()
    showToast(`Loaded ${t.rows.length} rows from the web.`)
  }
  return (
    <Dialog title="From Web" width={520} onCancel={onClose} okLabel={tables ? 'Load' : 'OK'} onOk={tables ? load : () => void fetchIt()}>
      <div style={{ fontSize: 12 }}>
        <label>URL<br />
          <input data-web-url autoFocus value={url} style={{ width: '100%', marginTop: 3 }} disabled={busy}
            onChange={(e) => { setUrl(e.target.value); setTables(null) }}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.stopPropagation(); void fetchIt() } }} />
        </label>
        {busy && <div style={{ marginTop: 8, color: '#605e5c' }}>Connecting…</div>}
        {err && <div style={{ marginTop: 8, color: '#a4262c' }}>{err}</div>}
        {tables && (
          <div style={{ marginTop: 8 }}>
            <div>Select a table:</div>
            <div style={{ border: '1px solid #d1d1d1', maxHeight: 130, overflowY: 'auto', marginTop: 3 }}>
              {tables.map((t, i) => (
                <div key={i} data-web-table={t.name} onClick={() => setPick(i)} style={{ padding: '2px 6px', cursor: 'pointer', background: i === pick ? '#c7e0f4' : undefined }}>{t.name} ({t.rows.length} rows)</div>
              ))}
            </div>
            <div style={{ marginTop: 8, overflow: 'auto', maxHeight: 150, border: '1px solid #d1d1d1' }}>
              <table style={{ borderCollapse: 'collapse', fontSize: 11 }}><tbody>
                {tables[pick].rows.slice(0, 8).map((r, i) => <tr key={i}>{r.slice(0, 8).map((v, j) => <td key={j} style={{ border: '1px solid #e1dfdd', padding: '1px 4px', whiteSpace: 'nowrap' }}>{v}</td>)}</tr>)}
              </tbody></table>
            </div>
          </div>
        )}
      </div>
    </Dialog>
  )
}

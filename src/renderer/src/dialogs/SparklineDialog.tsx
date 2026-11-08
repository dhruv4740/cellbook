// Insert > Sparklines > Line / Column / Win-Loss: pick the data and the cells that will hold the charts.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { rangeA1, SparklineDef } from '../model/types'
import { addSparklines } from '../data/sparklines'

const TITLE = { line: 'Line', column: 'Column', winloss: 'Win/Loss' }

export function SparklineDialog({ type, onClose }: { type: SparklineDef['type']; onClose: () => void }) {
  const sheet = store.active
  const sel = sheet.selection.ranges[sheet.selection.ranges.length - 1]
  const [data, setData] = useState(sel.r1 === sel.r2 && sel.c1 === sel.c2 ? '' : rangeA1(sel))
  const [loc, setLoc] = useState('')
  const [err, setErr] = useState<string | null>(null)
  const ok = () => {
    if (!store.allows(sheet, 'formatCells')) return
    const e = addSparklines(sheet, type, data, loc)
    if (e) setErr(e); else onClose()
  }
  return (
    <Dialog title="Create Sparklines" width={380} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 4 }}>Choose the data that you want ({TITLE[type]})</div>
        <div className="xl-field-row"><label style={{ width: 100 }}>Data Range:</label><input autoFocus value={data} onChange={(e) => setData(e.target.value)} style={{ flex: 1 }} placeholder="B2:G2" /></div>
        <div style={{ margin: '8px 0 4px' }}>Choose where you want the sparklines to be placed</div>
        <div className="xl-field-row"><label style={{ width: 100 }}>Location Range:</label><input value={loc} onChange={(e) => setLoc(e.target.value)} style={{ flex: 1 }} placeholder="H2" /></div>
        {err && <div style={{ color: '#a4262c', marginTop: 8 }}>{err}</div>}
      </div>
    </Dialog>
  )
}

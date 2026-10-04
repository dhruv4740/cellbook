// Data > Forecast Sheet.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { readForecastSource, createForecastSheet, forecastStep, parseForecastEnd } from '../data/forecast'

export function ForecastDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const src = readForecastSource(store, sheet, sheet.selection.ranges[sheet.selection.ranges.length - 1])
  const [chart, setChart] = useState<'line' | 'column'>('line')
  const [end, setEnd] = useState(() => (typeof src === 'string' ? '' : String(Math.round((src.xs[src.xs.length - 1] + forecastStep(src.xs) * Math.max(1, Math.ceil(src.xs.length / 2))) * 1e6) / 1e6)))
  const [ci, setCi] = useState(true)
  const [pct, setPct] = useState('95')
  const [season, setSeason] = useState<'auto' | 'manual'>('auto')
  const [period, setPeriod] = useState('1')
  const [err, setErr] = useState<string | null>(null)
  if (typeof src === 'string') {
    return (
      <Dialog title="Microsoft Excel" onCancel={onClose} cancelLabel="OK" width={380}>
        <div style={{ fontSize: 12 }}>{src}</div>
      </Dialog>
    )
  }
  const ok = () => {
    const e = parseForecastEnd(end)
    const p = Number(pct)
    if (e === null) return setErr('Enter Forecast End as a number or a date.')
    if (ci && !(p > 0 && p < 100)) return setErr('The confidence interval must be between 0 and 100.')
    const msg = createForecastSheet(store, src, { chart, end: e, confidence: ci ? p / 100 : null })
    if (msg) setErr(msg); else onClose()
  }
  return (
    <Dialog title="Create Forecast Worksheet" width={380} onCancel={onClose} onOk={ok} okLabel="Create">
      <div style={{ fontSize: 12 }}>
        <div>Chart type:</div>
        <label style={{ marginRight: 12 }}><input type="radio" checked={chart === 'line'} onChange={() => setChart('line')} /> Line</label>
        <label><input type="radio" checked={chart === 'column'} onChange={() => setChart('column')} /> Column</label>
        <div style={{ marginTop: 8 }}>Forecast End (number or date):</div>
        <input style={{ width: '100%' }} value={end} onChange={(e) => setEnd(e.target.value)} />
        <label style={{ display: 'block', marginTop: 8 }}>
          <input type="checkbox" checked={ci} onChange={(e) => setCi(e.target.checked)} /> Confidence Interval{' '}
          <input style={{ width: 44 }} disabled={!ci} value={pct} onChange={(e) => setPct(e.target.value)} /> %
        </label>
        <div style={{ marginTop: 8 }}>Seasonality:</div>
        <label style={{ marginRight: 12 }}><input type="radio" checked={season === 'auto'} onChange={() => setSeason('auto')} /> Detect automatically</label>
        <label><input type="radio" checked={season === 'manual'} onChange={() => setSeason('manual')} /> Set manually{' '}
          <input style={{ width: 44 }} disabled={season !== 'manual'} value={period} onChange={(e) => setPeriod(e.target.value)} /></label>
        <div style={{ color: '#666', marginTop: 6 }}>The forecast is a linear trend; seasonality is not modelled yet.</div>
        {err && <div style={{ color: '#a4262c', marginTop: 6 }}>{err}</div>}
      </div>
    </Dialog>
  )
}

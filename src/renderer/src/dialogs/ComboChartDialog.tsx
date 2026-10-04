// Insert > Combo Chart > Create Custom Combo Chart (and Chart Design > Change Chart Type > Combo): pick a chart type
// per series, and optionally put a series on the secondary axis.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { chartsOf, updateChart } from '../charts/chartsStore'
import { extractChartData } from '../charts/dataExtraction'
import { readerFor } from '../charts/chartsStore'
import { insertChartFromSelection, selectionSeriesNames } from '../charts/insertChart'
import { COMBO_GALLERY } from '../charts/gallery'
import { comboSeriesOf } from '../charts/chartCombo'

export function ComboChartDialog({ chartId, onClose }: { chartId?: string; onClose: () => void }) {
  const sheet = store.active
  const spec = chartId ? chartsOf(sheet).find((c) => c.id === chartId) : undefined
  const names = spec ? extractChartData(readerFor(store, sheet), spec.dataRange, spec.seriesInColumns, spec.firstIsCategories).series.map((s) => s.name) : selectionSeriesNames(store)
  const [rows, setRows] = useState(() => comboSeriesOf(spec ?? {}, names.length).map((r) => ({ type: r.type, secondary: !!r.secondary })))
  const set = (i: number, patch: Partial<(typeof rows)[number]>): void => setRows(rows.map((r, j) => (j === i ? { ...r, ...patch } : r)))
  const ok = (): void => {
    const comboSeries = rows.map((r) => ({ type: r.type, ...(r.secondary ? { secondary: true } : {}) }))
    if (spec) updateChart(store, sheet, spec.id, { type: 'combo', subtype: 'combo', comboSeries }, 'Change Chart Type')
    else insertChartFromSelection(store, COMBO_GALLERY[0], comboSeries)
    onClose()
  }
  return (
    <Dialog title={spec ? 'Change Chart Type' : 'Insert Chart'} width={420} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 6 }}>Choose the chart type and axis for your data series:</div>
        <table style={{ width: '100%', fontSize: 12 }}>
          <thead><tr><th style={{ textAlign: 'left' }}>Series Name</th><th style={{ textAlign: 'left' }}>Chart Type</th><th>Secondary Axis</th></tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i}>
                <td>{names[i] ?? `Series${i + 1}`}</td>
                <td>
                  <select value={r.type} onChange={(e) => set(i, { type: e.target.value as 'bar' | 'line' })}>
                    <option value="bar">Clustered Column</option>
                    <option value="line">Line</option>
                  </select>
                </td>
                <td style={{ textAlign: 'center' }}><input type="checkbox" checked={r.secondary} onChange={(e) => set(i, { secondary: e.target.checked })} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  )
}

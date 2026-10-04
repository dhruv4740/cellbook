// Cell right-click > Quick Analysis: one-click Formatting / Charts / Totals / Tables / Sparklines for the selection,
// each button reusing the existing conditional-format, chart, table, pivot and sparkline actions.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { rangeA1, normRange } from '../model/types'
import { setUi, showToast } from '../chrome/ui'
import { addCfRule, clearCfRules } from '../format/conditional'
import { COLUMN_BAR_GALLERY, LINE_AREA_GALLERY, PIE_DOUGHNUT_GALLERY, SCATTER_BUBBLE_GALLERY, GalleryEntry } from '../charts/gallery'
import { insertChartFromSelection } from '../charts/insertChart'
import { formatAsTable } from '../grid/commands'
import { addSparklines } from '../data/sparklines'
import { planTotals, TotalsFn, TotalsAxis } from '../data/quickAnalysis'
import { SparklineDef } from '../model/types'

type Tab = 'Formatting' | 'Charts' | 'Totals' | 'Tables' | 'Sparklines'
const TABS: Tab[] = ['Formatting', 'Charts', 'Totals', 'Tables', 'Sparklines']

interface Action { label: string; run: () => void }

export function QuickAnalysisDialog({ onClose }: { onClose: () => void }) {
  const [tab, setTab] = useState<Tab>('Formatting')
  const sheet = store.active
  const ranges = sheet.selection.ranges.map((r) => ({ ...r }))
  const last = normRange(ranges[ranges.length - 1])

  // Each action closes the popup first, then does its work (some open another dialog, which replaces this one).
  const act = (fn: () => void) => () => { onClose(); fn() }
  const chart = (e: GalleryEntry): Action => ({ label: e.label, run: () => insertChartFromSelection(store, e) })
  const total = (label: string, fn: TotalsFn, axis: TotalsAxis): Action => ({
    label,
    run: () => {
      if (!store.allows(sheet, 'formatCells')) return
      const cells = planTotals(last, fn, axis, (r, c) => typeof store.getValue(sheet, r, c) === 'number')
      if (!cells.length) { showToast('The selection has no numbers to total.'); return }
      store.setCells(sheet, cells, 'Quick Analysis Totals')
    }
  })
  const spark = (label: string, type: SparklineDef['type']): Action => ({
    label,
    run: () => {
      if (last.c1 === last.c2) { showToast('Select at least two columns of numbers for sparklines.'); return }
      const e = addSparklines(sheet, type, rangeA1(last), rangeA1({ r1: last.r1, c1: last.c2 + 1, r2: last.r2, c2: last.c2 + 1 }))
      if (e) setUi({ dialog: { kind: 'message', props: { text: e } } })
    }
  })

  const ACTIONS: Record<Tab, Action[]> = {
    Formatting: [
      { label: 'Data Bars', run: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#638EC6', gradient: true } }) },
      { label: 'Color Scale', run: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#F8696B', midColor: '#FFEB84', maxColor: '#63BE7B' } }) },
      { label: 'Icon Set', run: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '3TrafficLights1' } }) },
      { label: 'Greater Than...', run: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'greaterThan' } } }) },
      { label: 'Top 10%', run: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'top10Percent' } } }) },
      { label: 'Clear Format', run: () => clearCfRules(store, sheet, 'selection') }
    ],
    Charts: [
      chart(COLUMN_BAR_GALLERY[0]), chart(COLUMN_BAR_GALLERY[1]), chart(LINE_AREA_GALLERY[0]), chart(PIE_DOUGHNUT_GALLERY[0]), chart(SCATTER_BUBBLE_GALLERY[0])
    ],
    Totals: [
      total('Sum', 'SUM', 'col'), total('Average', 'AVERAGE', 'col'), total('Count', 'COUNT', 'col'),
      total('Sum (right)', 'SUM', 'row'), total('Average (right)', 'AVERAGE', 'row'), total('Count (right)', 'COUNT', 'row')
    ],
    Tables: [
      { label: 'Table', run: () => formatAsTable(0) },
      { label: 'PivotTable', run: () => setUi({ dialog: { kind: 'pivot' } }) }
    ],
    Sparklines: [spark('Line', 'line'), spark('Column', 'column'), spark('Win/Loss', 'winloss')]
  }

  return (
    <Dialog title="Quick Analysis" width={420} onCancel={onClose} hideFooter>
      <div style={{ display: 'flex', gap: 2, borderBottom: '1px solid #c8c8c8', marginBottom: 10 }}>
        {TABS.map((t) => (
          <button key={t} className={'xl-qa-tab' + (t === tab ? ' active' : '')} onClick={() => setTab(t)} style={{ border: 'none', background: t === tab ? '#fff' : 'transparent', borderBottom: t === tab ? '2px solid #217346' : '2px solid transparent', padding: '4px 8px', cursor: 'pointer', fontSize: 12 }}>{t}</button>
        ))}
      </div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, minHeight: 56 }}>
        {ACTIONS[tab].map((a) => (
          <button key={a.label} className="xl-qa-btn" onClick={act(a.run)} style={{ minWidth: 84, padding: '8px 6px', fontSize: 12, cursor: 'pointer' }}>{a.label}</button>
        ))}
      </div>
      <div style={{ fontSize: 11, color: '#666', marginTop: 8 }}>{tab === 'Totals' ? 'Totals are written below (or to the right of) the selection as formulas.' : tab === 'Sparklines' ? 'Sparklines go in the column to the right of the selection.' : ''}</div>
    </Dialog>
  )
}

// Excel's Page Setup dialog: Page / Margins / Header-Footer / Sheet tabs, opened from the Page
// Layout tab's dialog launchers (Page Setup, Margins, Print Titles) and from File > Print.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore, SheetState } from '../model/store'
import {
  PageSetup, PaperSize, PAPER_SIZES, MARGIN_PRESETS, HF_PRESETS, HeaderFooterText,
  rangeA1, parseA1, normRange, colName, colIndex
} from '../model/types'
import { HeaderFooterCustomDialog } from './HeaderFooterCustomDialog'
import './PageSetupDialog.css'

export type PageSetupTab = 'Page' | 'Margins' | 'Header/Footer' | 'Sheet'

export interface PageSetupDialogProps {
  store: WorkbookStore
  sheet: SheetState
  initialTab?: PageSetupTab
  onClose: () => void
  /** Called after OK with the finished PageSetup, in case the caller (Print pane) wants to react. */
  onApplied?: (ps: PageSetup) => void
}

function num(v: string, fallback: number): number {
  const n = Number(v)
  return Number.isFinite(n) ? n : fallback
}

export function PageSetupDialog({ store, sheet, initialTab, onClose, onApplied }: PageSetupDialogProps) {
  const [tab, setTab] = useState<PageSetupTab>(initialTab ?? 'Page')
  const [ps, setPs] = useState<PageSetup>(() => structuredClone(sheet.layout.pageSetup!))
  const [hfEditor, setHfEditor] = useState<null | { which: 'header' | 'footer' }>(null)

  const set = (patch: Partial<PageSetup>) => setPs((p) => ({ ...p, ...patch }))
  const setMargins = (patch: Partial<PageSetup['margins']>) => setPs((p) => ({ ...p, margins: { ...p.margins, ...patch } }))
  const setHF = (patch: Partial<PageSetup['headerFooter']>) => setPs((p) => ({ ...p, headerFooter: { ...p.headerFooter, ...patch } }))

  const apply = () => {
    store.updateLayout(sheet, 'Page Setup', (l) => { l.pageSetup = structuredClone(ps) })
    onApplied?.(ps)
  }
  const ok = () => { apply(); onClose() }

  const presetMatches = (preset: keyof typeof MARGIN_PRESETS): boolean => {
    const m = MARGIN_PRESETS[preset]
    return m.top === ps.margins.top && m.bottom === ps.margins.bottom && m.left === ps.margins.left && m.right === ps.margins.right
  }

  const hfPresetLabel = (t: HeaderFooterText): string => {
    const match = HF_PRESETS.find((p) => p.center === t.center && !t.left && !t.right)
    return match ? match.label : t.left || t.center || t.right ? '(custom)' : '(none)'
  }

  return (
    <Dialog title="Page Setup" width={520} onOk={ok} onCancel={onClose} okLabel="OK">
      <div className="xl-ps-root">
        <div className="xl-ps-tabs">
          {(['Page', 'Margins', 'Header/Footer', 'Sheet'] as PageSetupTab[]).map((t) => (
            <div key={t} className={'xl-ps-tab' + (tab === t ? ' active' : '')} onClick={() => setTab(t)}>{t}</div>
          ))}
        </div>
        <div className="xl-ps-body">
          {tab === 'Page' && (
            <div className="xl-ps-page">
              <fieldset className="xl-ps-fieldset">
                <legend>Orientation</legend>
                <label className="xl-ps-radio"><input type="radio" checked={ps.orientation === 'portrait'} onChange={() => set({ orientation: 'portrait' })} /> Portrait</label>
                <label className="xl-ps-radio"><input type="radio" checked={ps.orientation === 'landscape'} onChange={() => set({ orientation: 'landscape' })} /> Landscape</label>
              </fieldset>
              <fieldset className="xl-ps-fieldset">
                <legend>Scaling</legend>
                <label className="xl-ps-radio">
                  <input type="radio" checked={!ps.fitToPage} onChange={() => set({ fitToPage: false })} /> Adjust to:
                  <input className="xl-ps-num" type="number" min={10} max={400} value={ps.scalePercent}
                    onChange={(e) => set({ fitToPage: false, scalePercent: num(e.target.value, 100) })} /> % normal size
                </label>
                <label className="xl-ps-radio">
                  <input type="radio" checked={ps.fitToPage} onChange={() => set({ fitToPage: true })} /> Fit to:
                  <input className="xl-ps-num" type="number" min={0} max={99} value={ps.fitToWidth}
                    onChange={(e) => set({ fitToPage: true, fitToWidth: num(e.target.value, 1) })} /> page(s) wide by
                  <input className="xl-ps-num" type="number" min={0} max={99} value={ps.fitToHeight}
                    onChange={(e) => set({ fitToPage: true, fitToHeight: num(e.target.value, 1) })} /> tall
                </label>
              </fieldset>
              <div className="xl-ps-row">
                <label className="xl-ps-label">Paper size:</label>
                <select className="xl-ps-select" value={ps.paperSize} onChange={(e) => set({ paperSize: e.target.value as PaperSize })}>
                  {Object.entries(PAPER_SIZES).map(([k, v]) => <option key={k} value={k}>{v.label} ({v.w}&quot; x {v.h}&quot;)</option>)}
                </select>
              </div>
              <div className="xl-ps-row">
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.useFirstPageNumber} onChange={(e) => set({ useFirstPageNumber: e.target.checked })} /> First page number:</label>
                <input className="xl-ps-num" type="number" disabled={!ps.useFirstPageNumber} value={ps.firstPageNumber} onChange={(e) => set({ firstPageNumber: num(e.target.value, 1) })} />
              </div>
            </div>
          )}

          {tab === 'Margins' && (
            <div className="xl-ps-margins">
              <div className="xl-ps-margins-grid">
                {(['top', 'bottom', 'left', 'right', 'header', 'footer'] as const).map((k) => (
                  <label key={k} className="xl-ps-margin-field">
                    {k[0].toUpperCase() + k.slice(1)}:
                    <input className="xl-ps-num" type="number" step={0.05} min={0} value={ps.margins[k]} onChange={(e) => setMargins({ [k]: num(e.target.value, 0) } as any)} />
                  </label>
                ))}
              </div>
              <div className="xl-ps-margin-presets">
                <button className="xl-ttc-btn" onClick={() => setMargins(MARGIN_PRESETS.normal)}>Normal</button>
                <button className="xl-ttc-btn" onClick={() => setMargins(MARGIN_PRESETS.wide)}>Wide</button>
                <button className="xl-ttc-btn" onClick={() => setMargins(MARGIN_PRESETS.narrow)}>Narrow</button>
                <span className="xl-ps-preset-active">
                  {presetMatches('normal') ? 'Normal' : presetMatches('wide') ? 'Wide' : presetMatches('narrow') ? 'Narrow' : 'Custom'}
                </span>
              </div>
              <fieldset className="xl-ps-fieldset">
                <legend>Center on page</legend>
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.centerHorizontally} onChange={(e) => set({ centerHorizontally: e.target.checked })} /> Horizontally</label>
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.centerVertically} onChange={(e) => set({ centerVertically: e.target.checked })} /> Vertically</label>
              </fieldset>
            </div>
          )}

          {tab === 'Header/Footer' && (
            <div className="xl-ps-hf">
              <div className="xl-ps-row">
                <label className="xl-ps-label">Header:</label>
                <select className="xl-ps-select" value={hfPresetLabel(ps.headerFooter.oddHeader)}
                  onChange={(e) => { const p = HF_PRESETS.find((x) => x.label === e.target.value); if (p) setHF({ oddHeader: { left: '', center: p.center, right: '' } }) }}>
                  {HF_PRESETS.map((p) => <option key={p.label} value={p.label}>{p.label}</option>)}
                  <option value="(custom)">(custom)</option>
                </select>
                <button className="xl-ttc-btn" onClick={() => setHfEditor({ which: 'header' })}>Custom Header...</button>
              </div>
              <div className="xl-ps-row">
                <label className="xl-ps-label">Footer:</label>
                <select className="xl-ps-select" value={hfPresetLabel(ps.headerFooter.oddFooter)}
                  onChange={(e) => { const p = HF_PRESETS.find((x) => x.label === e.target.value); if (p) setHF({ oddFooter: { left: '', center: p.center, right: '' } }) }}>
                  {HF_PRESETS.map((p) => <option key={p.label} value={p.label}>{p.label}</option>)}
                  <option value="(custom)">(custom)</option>
                </select>
                <button className="xl-ttc-btn" onClick={() => setHfEditor({ which: 'footer' })}>Custom Footer...</button>
              </div>
              <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.headerFooter.differentOddEven} onChange={(e) => setHF({ differentOddEven: e.target.checked })} /> Different odd and even pages</label>
              <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.headerFooter.differentFirst} onChange={(e) => setHF({ differentFirst: e.target.checked })} /> Different first page</label>
              <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.headerFooter.scaleWithDoc} onChange={(e) => setHF({ scaleWithDoc: e.target.checked })} /> Scale with document</label>
              <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.headerFooter.alignWithMargins} onChange={(e) => setHF({ alignWithMargins: e.target.checked })} /> Align with page margins</label>
            </div>
          )}

          {tab === 'Sheet' && (
            <div className="xl-ps-sheet">
              <div className="xl-ps-row">
                <label className="xl-ps-label">Print area:</label>
                <input className="xl-ps-input" value={ps.printArea.map(rangeA1).join(',')}
                  onChange={(e) => set({ printArea: e.target.value.split(',').map((s) => parseA1(s.trim())).filter((r): r is NonNullable<typeof r> => !!r) })} />
                <button className="xl-ttc-btn" onClick={() => set({ printArea: [normRange(sheet.selection.ranges[0])] })}>Use Selection</button>
              </div>
              <fieldset className="xl-ps-fieldset">
                <legend>Print titles</legend>
                <div className="xl-ps-row">
                  <label className="xl-ps-label">Rows to repeat at top:</label>
                  <input className="xl-ps-input" value={ps.printTitleRows ? `$${ps.printTitleRows.r1 + 1}:$${ps.printTitleRows.r2 + 1}` : ''}
                    onChange={(e) => { const m = /(\d+):(\d+)/.exec(e.target.value); set({ printTitleRows: m ? { r1: +m[1] - 1, r2: +m[2] - 1 } : undefined }) }} />
                  <button className="xl-ttc-btn" onClick={() => { const r = normRange(sheet.selection.ranges[0]); set({ printTitleRows: { r1: r.r1, r2: r.r2 } }) }}>Use Rows</button>
                </div>
                <div className="xl-ps-row">
                  <label className="xl-ps-label">Columns to repeat at left:</label>
                  <input className="xl-ps-input" value={ps.printTitleCols ? `$${colName(ps.printTitleCols.c1)}:$${colName(ps.printTitleCols.c2)}` : ''}
                    onChange={(e) => { const m = /([A-Z]+):([A-Z]+)/i.exec(e.target.value); set({ printTitleCols: m ? { c1: colIndex(m[1]), c2: colIndex(m[2]) } : undefined }) }} />
                  <button className="xl-ttc-btn" onClick={() => { const r = normRange(sheet.selection.ranges[0]); set({ printTitleCols: { c1: r.c1, c2: r.c2 } }) }}>Use Columns</button>
                </div>
              </fieldset>
              <fieldset className="xl-ps-fieldset">
                <legend>Print</legend>
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> Gridlines</label>
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.headings} onChange={(e) => set({ headings: e.target.checked })} /> Row and column headings</label>
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.blackAndWhite} onChange={(e) => set({ blackAndWhite: e.target.checked })} /> Black and white</label>
                <label className="xl-ps-checkbox"><input type="checkbox" checked={ps.draft} onChange={(e) => set({ draft: e.target.checked })} /> Draft quality</label>
              </fieldset>
              <fieldset className="xl-ps-fieldset">
                <legend>Page order</legend>
                <label className="xl-ps-radio"><input type="radio" checked={ps.pageOrder === 'downThenOver'} onChange={() => set({ pageOrder: 'downThenOver' })} /> Down, then over</label>
                <label className="xl-ps-radio"><input type="radio" checked={ps.pageOrder === 'overThenDown'} onChange={() => set({ pageOrder: 'overThenDown' })} /> Over, then down</label>
              </fieldset>
            </div>
          )}
        </div>
      </div>
      {hfEditor && (
        <HeaderFooterCustomDialog
          title={hfEditor.which === 'header' ? 'Header' : 'Footer'}
          initial={hfEditor.which === 'header' ? ps.headerFooter.oddHeader : ps.headerFooter.oddFooter}
          onOk={(v) => { setHF(hfEditor.which === 'header' ? { oddHeader: v } : { oddFooter: v }); setHfEditor(null) }}
          onCancel={() => setHfEditor(null)}
        />
      )}
    </Dialog>
  )
}

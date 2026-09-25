// Excel 365's File backstage view: full-window overlay with the green nav pane on the left.
import { useEffect, useState } from 'react'
import { store, useStore } from '../model/store'
import { setUi, useUi } from '../chrome/ui'
import {
  openFile, openRecent, save, saveAs, newWorkbook, closeWorkbook, downloadCsv,
  getRecentFiles, installFileShortcuts, writeActiveCsv, type RecentFile
} from './fileOps'
import {
  HomeRegular, DocumentAdd20Regular, FolderOpen20Regular, LockClosed20Regular,
  CheckmarkCircle20Regular, ArrowClockwise20Regular, Folder20Regular, DocumentPdf20Regular,
  ArrowExportUp20Regular, ChevronLeft20Regular
} from '@fluentui/react-icons'
import { PageSetup, PaperSize, PAPER_SIZES, MARGIN_PRESETS, defaultPageSetup, normRange } from '../model/types'
import { buildSinglePageDoc, computePrintPlan, printTargets, exportTargetsToPdf, PrintTarget } from '../print'
import type { PrinterInfo } from '../../../preload'
import './Backstage.css'

// Installing the global Ctrl+N/O/S/F12/Ctrl+W shortcuts and the Electron close-confirm listener happens once,
// as soon as this module loads — App.tsx imports Backstage unconditionally, so this always runs at startup.
installFileShortcuts()

type Tab =
  | 'home' | 'new' | 'open' | 'info' | 'save' | 'saveas' | 'print' | 'share' | 'export' | 'publish' | 'close'

const TEMPLATES = ['Budget', 'Calendar', 'Invoice', 'Timesheet', 'Planner']

function greeting(): string {
  const h = new Date().getHours()
  if (h < 12) return 'Good morning'
  if (h < 18) return 'Good afternoon'
  return 'Good evening'
}

function close(): void {
  setUi({ backstage: false })
}

export default function Backstage() {
  useStore()
  const ui = useUi()
  const [tab, setTab] = useState<Tab>('home')

  useEffect(() => {
    // Alt,F then a nav-item letter (KeyTips.tsx exits its own overlay the instant the File tab
    // opens the backstage, since this view has no ribbon-style badges of its own) - Excel 2021's
    // real backstage access keys: N New, O Open, I Info, S Save, A Save As, P Print, E Export,
    // C Close. The single Alt keypress that opened the backstage still has its keyup pending, so
    // this only reacts to the letter itself (no Alt required), matching what a user types next.
    const NAV_KEYS: Record<string, () => void> = {
      N: () => go('new'), O: () => go('open'), I: () => go('info'), S: () => void save(),
      A: () => go('saveas'), P: () => go('print'), E: () => go('export'), C: () => void closeWorkbook().then((ok) => ok && close())
    }
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') { close(); return }
      if (e.ctrlKey || e.metaKey || e.altKey || e.key.length !== 1) return
      const fn = NAV_KEYS[e.key.toUpperCase()]
      if (fn && !(e.target instanceof HTMLInputElement) && !(e.target instanceof HTMLTextAreaElement)) {
        e.preventDefault()
        fn()
      }
    }
    window.addEventListener('keydown', key)
    return () => window.removeEventListener('keydown', key)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Excel always opens the backstage on the Home page, however it was left last time.
  useEffect(() => {
    if (ui.backstage) {
      setTab((ui.backstageTab as Tab) ?? 'home')
      if (ui.backstageTab) setUi({ backstageTab: null })
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ui.backstage])

  if (!ui.backstage) return null

  const go = (t: Tab): void => setTab(t)

  return (
    <div className="bs-root">
      <nav className="bs-nav">
        <div className="bs-back" onClick={close} title="Back">
          <ChevronLeft20Regular />
        </div>
        <div className={`bs-item primary${tab === 'home' ? ' active' : ''}`} onClick={() => go('home')}><HomeRegular fontSize={18} /><span>Home</span></div>
        <div className={`bs-item primary${tab === 'new' ? ' active' : ''}`} onClick={() => go('new')}><DocumentAdd20Regular /><span>New</span></div>
        <div className={`bs-item primary${tab === 'open' ? ' active' : ''}`} onClick={() => go('open')}><FolderOpen20Regular /><span>Open</span></div>
        <div className="bs-sep" />
        <div className={`bs-item${tab === 'info' ? ' active' : ''}`} onClick={() => go('info')}>Info</div>
        <div className="bs-item" onClick={() => { void save() }}>Save</div>
        <div className={`bs-item${tab === 'saveas' ? ' active' : ''}`} onClick={() => go('saveas')}>Save As</div>
        <div className={`bs-item${tab === 'print' ? ' active' : ''}`} onClick={() => go('print')}>Print</div>
        <div className={`bs-item${tab === 'share' ? ' active' : ''}`} onClick={() => go('share')}>Share</div>
        <div className={`bs-item${tab === 'export' ? ' active' : ''}`} onClick={() => go('export')}>Export</div>
        <div className={`bs-item${tab === 'publish' ? ' active' : ''}`} onClick={() => go('publish')}>Publish</div>
        <div className="bs-item" onClick={() => { void closeWorkbook().then((ok) => ok && close()) }}>Close</div>
        <div className="bs-spacer" />
        <div className="bs-bottom">
          <div className="bs-item">Account</div>
          <div className="bs-item">Feedback</div>
          <div className="bs-item">Options</div>
        </div>
      </nav>
      <div className="bs-content">
        {tab === 'home' && <HomePane />}
        {tab === 'new' && <NewPane />}
        {tab === 'open' && <PlacesPane mode="open" />}
        {tab === 'info' && <InfoPane />}
        {tab === 'saveas' && <PlacesPane mode="saveas" />}
        {tab === 'print' && <PrintPane />}
        {tab === 'share' && <SimplePane title="Share" text="Save this workbook to the cloud to share it with other people and collaborate at the same time." />}
        {tab === 'export' && <ExportPane />}
        {tab === 'publish' && <SimplePane title="Publish" text="Publish this workbook to Power BI to explore and visualize your data." />}
      </div>
    </div>
  )
}

function TemplateGrid({ onPick }: { onPick: () => void }) {
  return (
    <div className="bs-templates">
      <div className="bs-template" onClick={onPick}>
        <div className="bs-template-thumb blank">
          <span className="bs-template-plus">+</span>
        </div>
        <div className="bs-template-name">Blank workbook</div>
      </div>
      {TEMPLATES.map((t) => (
        <div className="bs-template" key={t} onClick={onPick}>
          <div className="bs-template-thumb" style={{ borderTop: '3px solid #217346' }}>{t}</div>
          <div className="bs-template-name">{t}</div>
        </div>
      ))}
    </div>
  )
}

function RecentList({ onOpen }: { onOpen: (f: RecentFile) => void }) {
  const recents = getRecentFiles()
  if (!recents.length) return <div className="bs-empty">No recent workbooks yet.</div>
  return (
    <div className="bs-recent-list">
      {recents.map((f, i) => (
        <div className="bs-recent-row" key={i} onClick={() => onOpen(f)}>
          <span className="bs-recent-icon"><DocumentAdd20Regular fontSize={16} /></span>
          <span className="bs-recent-name">{f.name}</span>
          <span className="bs-recent-path">{f.path ?? ''}</span>
        </div>
      ))}
    </div>
  )
}

function HomePane() {
  return (
    <div className="bs-pad">
      <div className="bs-greeting">{greeting()}</div>
      <div className="bs-h2">New</div>
      <TemplateGrid onPick={() => { void newWorkbook().then((ok) => ok && close()) }} />
      <div className="bs-h2">Recent</div>
      <RecentList onOpen={(f) => { void openRecent(f).then((ok) => ok && close()) }} />
    </div>
  )
}

function NewPane() {
  return (
    <div className="bs-pad">
      <div className="bs-places-title" style={{ padding: 0, marginBottom: 8 }}>New</div>
      <input className="bs-search" placeholder="Search for online templates" />
      <div className="bs-h2">Blank workbook &amp; templates</div>
      <TemplateGrid onPick={() => { void newWorkbook().then((ok) => ok && close()) }} />
    </div>
  )
}

function PlacesPane({ mode }: { mode: 'open' | 'saveas' }) {
  const [place, setPlace] = useState<'recent' | 'this-pc' | 'onedrive' | 'shared'>(mode === 'open' ? 'recent' : 'this-pc')
  const browse = (): void => {
    if (mode === 'open') void openFile().then((ok) => ok && close())
    else void saveAs().then((ok) => ok && close())
  }
  return (
    <div className="bs-places-layout">
      <div className="bs-places-col">
        {mode === 'open' && (
          <div className={`bs-place-row${place === 'recent' ? ' active' : ''}`} onClick={() => setPlace('recent')}>Recent</div>
        )}
        {mode === 'open' && (
          <div className={`bs-place-row${place === 'shared' ? ' active' : ''}`} onClick={() => setPlace('shared')}>Shared with Me</div>
        )}
        <div className={`bs-place-row${place === 'onedrive' ? ' active' : ''}`} onClick={() => setPlace('onedrive')}>
          OneDrive - Personal
        </div>
        <div className={`bs-place-row${place === 'this-pc' ? ' active' : ''}`} onClick={() => setPlace('this-pc')}>This PC</div>
        <div className="bs-place-row" onClick={browse}>+ Add a Place</div>
        <button className="bs-btn browse-btn bs-browse-btn" onClick={browse}>Browse</button>
      </div>
      <div className="bs-pad" style={{ flex: 1 }}>
        <div className="bs-places-title" style={{ padding: 0 }}>{mode === 'open' ? 'Open' : 'Save As'}</div>
        {place === 'recent' && mode === 'open' && <RecentList onOpen={(f) => { void openRecent(f).then((ok) => ok && close()) }} />}
        {place !== 'recent' && (
          <div className="bs-empty">
            {place === 'this-pc' ? 'Browse for a file on this computer.' : 'Connect a cloud account to see files here.'}
            <div style={{ marginTop: 16 }}>
              <button className="bs-btn" onClick={browse}>Browse</button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

function InfoPane() {
  const used = store.usedRange(store.active)
  const cellCount = used ? (used.r2 + 1) * (used.c2 + 1) : 0
  return (
    <div className="bs-pad">
      <div className="bs-info-header">
        <div>
          <div className="bs-info-title">{store.fileName}</div>
          <div className="bs-info-path">{store.filePath ?? 'This PC › Documents'}</div>
        </div>
      </div>
      <div className="bs-tiles">
        <div className="bs-tile">
          <div className="bs-tile-icon"><LockClosed20Regular fontSize={26} /></div>
          <div>
            <div className="bs-tile-title">Protect Workbook</div>
            <div className="bs-tile-desc">Control what types of changes people can make to this workbook.</div>
          </div>
        </div>
        <div className="bs-tile">
          <div className="bs-tile-icon"><CheckmarkCircle20Regular fontSize={26} /></div>
          <div>
            <div className="bs-tile-title">Inspect Workbook</div>
            <div className="bs-tile-desc">Check the workbook for hidden properties or personal information.</div>
          </div>
        </div>
        <div className="bs-tile">
          <div className="bs-tile-icon"><ArrowClockwise20Regular fontSize={26} /></div>
          <div>
            <div className="bs-tile-title">Version History</div>
            <div className="bs-tile-desc">View and restore previous versions of this file.</div>
          </div>
        </div>
        <div className="bs-tile">
          <div className="bs-tile-icon"><Folder20Regular fontSize={26} /></div>
          <div>
            <div className="bs-tile-title">Manage Workbook</div>
            <div className="bs-tile-desc">Recover unsaved workbooks and manage checked-out files.</div>
          </div>
        </div>
      </div>
      <div className="bs-h2">Properties</div>
      <div className="bs-props">
        <div className="bs-prop-row"><span className="bs-prop-label">Size</span><span>{(cellCount * 12).toLocaleString()} bytes</span></div>
        <div className="bs-prop-row"><span className="bs-prop-label">Title</span><span>{store.fileName}</span></div>
        <div className="bs-prop-row"><span className="bs-prop-label">Tags</span><span>Add a tag</span></div>
        <div className="bs-prop-row"><span className="bs-prop-label">Sheets</span><span>{store.sheets.length}</span></div>
        <div className="bs-prop-row"><span className="bs-prop-label">Last Modified</span><span>{new Date().toLocaleString()}</span></div>
        <div className="bs-prop-row"><span className="bs-prop-label">Created</span><span>{new Date().toLocaleString()}</span></div>
        <div className="bs-prop-row"><span className="bs-prop-label">Author</span><span>{'​'}</span></div>
      </div>
    </div>
  )
}

function PrintPane() {
  useStore()
  const sheet = store.active
  const ps = sheet.layout.pageSetup ?? defaultPageSetup()
  const [pageIdx, setPageIdx] = useState(0)
  const [zoom, setZoom] = useState<'fit' | 'full'>('fit')
  const [copies, setCopies] = useState(1)
  const [printWhat, setPrintWhat] = useState<'active' | 'workbook' | 'selection'>('active')
  const [printers, setPrinters] = useState<PrinterInfo[]>([])
  const [printer, setPrinter] = useState('')
  const [printing, setPrinting] = useState(false)

  useEffect(() => {
    void window.api?.getPrinters?.().then((list) => {
      setPrinters(list)
      if (list.length && !printer) setPrinter(list.find((p) => p.isDefault)?.name ?? list[0].name)
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const patch = (fn: (p: PageSetup) => void) => store.updateLayout(sheet, 'Page Setup', (l) => {
    const next = structuredClone(l.pageSetup ?? defaultPageSetup())
    fn(next)
    l.pageSetup = next
  })

  const targets = (): PrintTarget[] =>
    printWhat === 'workbook' ? store.sheets.map((s) => ({ sheet: s, pageSetup: s.layout.pageSetup ?? defaultPageSetup() })) : [{ sheet, pageSetup: ps }]

  const doc = buildSinglePageDoc(store, sheet, ps, pageIdx, store.fileName)
  const pageCount = printWhat === 'workbook'
    ? targets().reduce((n, t) => n + computePrintPlan(store, t.sheet, t.pageSetup).pages.length, 0)
    : doc.pageCount

  const doPrint = async () => {
    setPrinting(true)
    try {
      await printTargets(store, targets(), store.fileName, {
        copies, collate: true, deviceName: printer, landscape: ps.orientation === 'landscape',
        paperWidthPx: doc.widthPx, paperHeightPx: doc.heightPx
      })
    } finally {
      setPrinting(false)
    }
  }

  return (
    <div className="bs-print-layout">
      <div className="bs-print-settings">
        <button className="bs-btn" style={{ width: '100%', marginBottom: 20 }} onClick={doPrint} disabled={printing}>
          {printing ? 'Printing...' : 'Print'}
        </button>
        <div className="bs-print-label">Copies</div>
        <input className="bs-select" value={copies} type="number" min={1} onChange={(e) => setCopies(Math.max(1, Number(e.target.value) || 1))} />
        <div className="bs-print-label">Printer</div>
        <select className="bs-select" value={printer} onChange={(e) => setPrinter(e.target.value)}>
          {printers.length === 0 && <option value="">Microsoft Print to PDF</option>}
          {printers.map((p) => <option key={p.name} value={p.name}>{p.displayName}{p.isDefault ? ' (default)' : ''}</option>)}
        </select>
        <div className="bs-print-label">Settings</div>
        <select className="bs-select" value={printWhat} onChange={(e) => setPrintWhat(e.target.value as typeof printWhat)}>
          <option value="active">Print Active Sheets</option>
          <option value="workbook">Print Entire Workbook</option>
          <option value="selection">Print Selection</option>
        </select>
        {printWhat === 'selection' && (
          <button className="bs-btn" style={{ marginTop: 4 }} onClick={() => patch((p) => (p.printArea = sheet.selection.ranges.map(normRange)))}>
            Use Current Selection as Print Area
          </button>
        )}
        <div className="bs-print-pages-inline">
          <input className="bs-select" style={{ width: 50 }} placeholder="1" />
          <span>to</span>
          <input className="bs-select" style={{ width: 50 }} placeholder={String(pageCount)} />
          <span className="bs-print-label" style={{ margin: 0 }}>Pages</span>
        </div>

        <select className="bs-select" value={ps.orientation} onChange={(e) => patch((p) => (p.orientation = e.target.value as 'portrait' | 'landscape'))}>
          <option value="portrait">Portrait Orientation</option>
          <option value="landscape">Landscape Orientation</option>
        </select>
        <select className="bs-select" value={ps.paperSize} onChange={(e) => patch((p) => (p.paperSize = e.target.value as PaperSize))}>
          {Object.entries(PAPER_SIZES).map(([k, v]) => <option key={k} value={k}>{v.label} ({v.w}&quot; x {v.h}&quot;)</option>)}
        </select>
        <select
          className="bs-select"
          value={
            ps.margins.top === MARGIN_PRESETS.normal.top && ps.margins.left === MARGIN_PRESETS.normal.left ? 'normal'
              : ps.margins.top === MARGIN_PRESETS.wide.top && ps.margins.left === MARGIN_PRESETS.wide.left ? 'wide'
                : ps.margins.top === MARGIN_PRESETS.narrow.top && ps.margins.left === MARGIN_PRESETS.narrow.left ? 'narrow' : 'custom'
          }
          onChange={(e) => { const k = e.target.value; if (k !== 'custom') patch((p) => (p.margins = { ...MARGIN_PRESETS[k as 'normal' | 'wide' | 'narrow'] })) }}
        >
          <option value="normal">Normal Margins</option>
          <option value="wide">Wide Margins</option>
          <option value="narrow">Narrow Margins</option>
          <option value="custom">Custom Margins</option>
        </select>
        <select
          className="bs-select"
          value={!ps.fitToPage ? 'noscale' : ps.fitToWidth === 1 && ps.fitToHeight === 1 ? 'fitsheet' : ps.fitToWidth === 1 && ps.fitToHeight === 0 ? 'fitcols' : ps.fitToHeight === 1 && ps.fitToWidth === 0 ? 'fitrows' : 'noscale'}
          onChange={(e) => {
            const v = e.target.value
            if (v === 'noscale') patch((p) => { p.fitToPage = false; p.scalePercent = 100 })
            else if (v === 'fitsheet') patch((p) => { p.fitToPage = true; p.fitToWidth = 1; p.fitToHeight = 1 })
            else if (v === 'fitcols') patch((p) => { p.fitToPage = true; p.fitToWidth = 1; p.fitToHeight = 0 })
            else if (v === 'fitrows') patch((p) => { p.fitToPage = true; p.fitToWidth = 0; p.fitToHeight = 1 })
          }}
        >
          <option value="noscale">No Scaling</option>
          <option value="fitsheet">Fit Sheet on One Page</option>
          <option value="fitcols">Fit All Columns on One Page</option>
          <option value="fitrows">Fit All Rows on One Page</option>
        </select>
        <button className="bs-btn" style={{ width: '100%', marginTop: 12 }} onClick={() => setUi({ dialog: { kind: 'pageSetup', props: { tab: 'Page' } } })}>
          Page Setup
        </button>
      </div>
      <div className="bs-print-preview-col">
        <div className="bs-print-preview">
          <iframe key={pageIdx + ':' + JSON.stringify(ps)} className={'bs-print-frame' + (zoom === 'fit' ? ' fit' : '')} srcDoc={doc.html} title="Print preview" />
        </div>
        <div className="bs-print-preview-bar">
          <button className="bs-print-nav" disabled={pageIdx <= 0} onClick={() => setPageIdx((i) => Math.max(0, i - 1))}>◀</button>
          <span>Page {Math.min(pageIdx + 1, pageCount)} of {pageCount}</span>
          <button className="bs-print-nav" disabled={pageIdx >= pageCount - 1} onClick={() => setPageIdx((i) => Math.min(pageCount - 1, i + 1))}>▶</button>
          <span className="bs-print-bar-sep" />
          <button className={'bs-print-nav' + (zoom === 'fit' ? ' active' : '')} onClick={() => setZoom('fit')}>Zoom to Page</button>
          <button className={'bs-print-nav' + (zoom === 'full' ? ' active' : '')} onClick={() => setZoom('full')}>Show Margins</button>
        </div>
      </div>
    </div>
  )
}

/** Export > Create PDF: webContents.printToPDF (Electron) via print/printActions.ts, saved through
 *  a normal browser download; falls back to window.print() ("Save as PDF" in the OS print dialog)
 *  when there's no Electron bridge (npm run web). */
async function createPdf(): Promise<void> {
  const sheet = store.active
  const ps = sheet.layout.pageSetup ?? defaultPageSetup()
  const bytes = await exportTargetsToPdf(store, [{ sheet, pageSetup: ps }], store.fileName)
  if (!bytes) return // browser fallback already invoked window.print()
  const blob = new Blob([bytes as BlobPart], { type: 'application/pdf' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = store.fileName.replace(/\.xlsx?$/i, '') + '.pdf'
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 2000)
}

function ExportPane() {
  return (
    <div className="bs-pad">
      <div className="bs-places-title" style={{ padding: 0 }}>Export</div>
      <div className="bs-export-tile" onClick={() => { void createPdf() }}>
        <div className="bs-tile-icon"><DocumentPdf20Regular fontSize={26} /></div>
        <div>
          <div className="bs-tile-title">Create PDF/XPS Document</div>
          <div className="bs-tile-desc">Save a copy of the workbook that's fixed-layout and easy to share (Microsoft Print to PDF).</div>
        </div>
      </div>
      <div className="bs-export-tile" onClick={downloadCsv}>
        <div className="bs-tile-icon"><ArrowExportUp20Regular fontSize={26} /></div>
        <div>
          <div className="bs-tile-title">Change File Type</div>
          <div className="bs-tile-desc">Export the active sheet as CSV (Comma delimited) — {writeActiveCsv().length.toLocaleString()} characters.</div>
        </div>
      </div>
    </div>
  )
}

function SimplePane({ title, text }: { title: string; text: string }) {
  return (
    <div className="bs-pad">
      <div className="bs-places-title" style={{ padding: 0 }}>{title}</div>
      <div className="bs-empty" style={{ maxWidth: 480 }}>{text}</div>
    </div>
  )
}

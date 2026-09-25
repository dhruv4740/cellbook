// Excel 365's File backstage view: full-window overlay with the green nav pane on the left.
import { useEffect, useMemo, useState } from 'react'
import { buildPrintHtml, runPrint, PrintOptions, DEFAULT_PRINT, optionsFor } from './print'
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
  const [o, setO] = useState<PrintOptions>(() => optionsFor(store.active))
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState<string | null>(null)
  const html = useMemo(() => buildPrintHtml(o), [o])
  const set = (p: Partial<PrintOptions>) => {
    const next = { ...o, ...p }
    setO(next)
    // Page settings belong to the sheet (and are saved with it), as in Excel.
    if (!('what' in p)) {
      store.updateLayout(store.active, 'Page Setup', (l) => {
        l.pageSetup = { landscape: next.landscape, pageSize: next.pageSize, margins: next.margins, fitColumns: next.scaling === 'fitColumns', gridlines: next.gridlines, headings: next.headings }
      })
    }
  }
  const go = async (mode: 'print' | 'pdf') => {
    setBusy(true); setErr(null)
    const e = await runPrint(o, mode)
    setBusy(false)
    if (e) setErr(e)
  }
  return (
    <div className="bs-print-layout">
      <div className="bs-print-settings">
        <button className="bs-btn" style={{ width: '100%', marginBottom: 8 }} disabled={busy || !html} onClick={() => void go('print')}>Print</button>
        <button className="bs-btn" style={{ width: '100%', marginBottom: 20 }} disabled={busy || !html} onClick={() => void go('pdf')}>Save as PDF</button>
        {err && <div style={{ color: '#a4262c', fontSize: 12, marginBottom: 10 }}>{err}</div>}
        <div className="bs-print-label">Settings</div>
        <select className="bs-select" value={o.what} onChange={(e) => set({ what: e.target.value as PrintOptions['what'] })}>
          <option value="active">Print Active Sheets</option>
          <option value="workbook">Print Entire Workbook</option>
          <option value="selection">Print Selection</option>
        </select>
        <select className="bs-select" value={o.landscape ? 'landscape' : 'portrait'} onChange={(e) => set({ landscape: e.target.value === 'landscape' })}>
          <option value="portrait">Portrait Orientation</option>
          <option value="landscape">Landscape Orientation</option>
        </select>
        <select className="bs-select" value={o.pageSize} onChange={(e) => set({ pageSize: e.target.value as PrintOptions['pageSize'] })}>
          <option value="A4">A4 (21 cm x 29.7 cm)</option>
          <option value="Letter">Letter (21.59 cm x 27.94 cm)</option>
          <option value="Legal">Legal (21.59 cm x 35.56 cm)</option>
        </select>
        <select className="bs-select" value={o.margins} onChange={(e) => set({ margins: e.target.value as PrintOptions['margins'] })}>
          <option value="normal">Normal Margins</option>
          <option value="wide">Wide Margins</option>
          <option value="narrow">Narrow Margins</option>
        </select>
        <select className="bs-select" value={o.scaling} onChange={(e) => set({ scaling: e.target.value as PrintOptions['scaling'] })}>
          <option value="none">No Scaling</option>
          <option value="fitColumns">Fit All Columns on One Page</option>
        </select>
        <label style={{ display: 'block', fontSize: 12, marginTop: 8 }}><input type="checkbox" checked={o.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> Print gridlines</label>
        <label style={{ display: 'block', fontSize: 12 }}><input type="checkbox" checked={o.headings} onChange={(e) => set({ headings: e.target.checked })} /> Print row and column headings</label>
        <div style={{ fontSize: 11, color: '#666', marginTop: 12 }}>The printer, copies and page range are chosen in the Windows print dialog.</div>
      </div>
      <div className="bs-print-preview">
        {html
          ? <iframe title="Print preview" srcDoc={html} className="bs-print-frame" style={{ width: (o.landscape ? 11.69 : 8.27) * 96 * 0.8, height: '100%', border: '1px solid #ccc', background: '#fff' }} />
          : <div className="bs-print-page"><div style={{ padding: 20, color: '#666' }}>There is nothing to print on this sheet.</div></div>}
      </div>
    </div>
  )
}

function ExportPane() {
  return (
    <div className="bs-pad">
      <div className="bs-places-title" style={{ padding: 0 }}>Export</div>
      <div className="bs-export-tile" onClick={() => void runPrint(optionsFor(store.active), 'pdf')}>
        <div className="bs-tile-icon"><DocumentPdf20Regular fontSize={26} /></div>
        <div>
          <div className="bs-tile-title">Create PDF/XPS Document</div>
          <div className="bs-tile-desc">Save a copy of the workbook that's fixed-layout and easy to share, as a PDF file.</div>
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

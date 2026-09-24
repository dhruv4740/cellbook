// Excel 365's File backstage view: full-window overlay with the green nav pane on the left.
import { useEffect, useState } from 'react'
import { store, useStore } from '../model/store'
import { setUi, useUi } from '../chrome/ui'
import {
  openFile, save, saveAs, newWorkbook, closeWorkbook, downloadCsv,
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
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close()
    }
    window.addEventListener('keydown', key)
    return () => window.removeEventListener('keydown', key)
  }, [])

  // Excel always opens the backstage on the Home page, however it was left last time.
  useEffect(() => {
    if (ui.backstage) setTab('home')
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
        <div className="bs-item" onClick={() => { void closeWorkbook(); close() }}>Close</div>
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
      <TemplateGrid onPick={() => { void newWorkbook(); close() }} />
      <div className="bs-h2">Recent</div>
      <RecentList onOpen={() => { void openFile(); close() }} />
    </div>
  )
}

function NewPane() {
  return (
    <div className="bs-pad">
      <div className="bs-places-title" style={{ padding: 0, marginBottom: 8 }}>New</div>
      <input className="bs-search" placeholder="Search for online templates" />
      <div className="bs-h2">Blank workbook &amp; templates</div>
      <TemplateGrid onPick={() => { void newWorkbook(); close() }} />
    </div>
  )
}

function PlacesPane({ mode }: { mode: 'open' | 'saveas' }) {
  const [place, setPlace] = useState<'recent' | 'this-pc' | 'onedrive' | 'shared'>(mode === 'open' ? 'recent' : 'this-pc')
  const browse = (): void => {
    if (mode === 'open') void openFile().then(close)
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
        {place === 'recent' && mode === 'open' && <RecentList onOpen={() => { void openFile().then(close) }} />}
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
  const sheet = store.active
  const used = store.usedRange(sheet)
  const rows = used ? Math.min(used.r2 + 1, 60) : 0
  const cols = used ? Math.min(used.c2 + 1, 20) : 0
  return (
    <div className="bs-print-layout">
      <div className="bs-print-settings">
        <button className="bs-btn" style={{ width: '100%', marginBottom: 20 }} onClick={() => window.print()}>Print</button>
        <div className="bs-print-label">Copies</div>
        <input className="bs-select" defaultValue={1} type="number" min={1} />
        <div className="bs-print-label">Printer</div>
        <select className="bs-select" defaultValue="pdf">
          <option value="pdf">Microsoft Print to PDF</option>
          <option value="xps">Microsoft XPS Document Writer</option>
        </select>
        <div className="bs-print-label">Settings</div>
        <select className="bs-select" defaultValue="active">
          <option value="active">Print Active Sheets</option>
          <option value="workbook">Print Entire Workbook</option>
          <option value="selection">Print Selection</option>
        </select>
        <select className="bs-select" defaultValue="portrait">
          <option value="portrait">Portrait Orientation</option>
          <option value="landscape">Landscape Orientation</option>
        </select>
        <select className="bs-select" defaultValue="letter">
          <option value="letter">Letter (8.5&quot; x 11&quot;)</option>
          <option value="a4">A4 (8.27&quot; x 11.69&quot;)</option>
        </select>
        <select className="bs-select" defaultValue="normal">
          <option value="normal">Normal Margins</option>
          <option value="wide">Wide Margins</option>
          <option value="narrow">Narrow Margins</option>
        </select>
        <select className="bs-select" defaultValue="noscale">
          <option value="noscale">No Scaling</option>
          <option value="fitsheet">Fit Sheet on One Page</option>
        </select>
      </div>
      <div className="bs-print-preview">
        <div className="bs-print-page">
          <table>
            <tbody>
              {Array.from({ length: rows }).map((_, r) => (
                <tr key={r}>
                  {Array.from({ length: cols }).map((_, c) => (
                    <td key={c}>{store.getDisplay(sheet, r, c).text}</td>
                  ))}
                </tr>
              ))}
              {rows === 0 && <tr><td>&nbsp;</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

function ExportPane() {
  return (
    <div className="bs-pad">
      <div className="bs-places-title" style={{ padding: 0 }}>Export</div>
      <div className="bs-export-tile" onClick={() => window.print()}>
        <div className="bs-tile-icon"><DocumentPdf20Regular fontSize={26} /></div>
        <div>
          <div className="bs-tile-title">Create PDF/XPS Document</div>
          <div className="bs-tile-desc">Save a copy of the workbook that's fixed-layout and easy to share, using the browser's Print to PDF.</div>
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

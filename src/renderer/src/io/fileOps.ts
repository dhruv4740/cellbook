// File menu operations: New, Open, Save, Save As, Close — Electron-backed, with a browser fallback.
import { store } from '../model/store'
import { readXlsx, writeXlsx } from './xlsx'
import { readCsv, writeCsv } from './csv'
import { confirmSaveChanges } from './saveConfirm'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const CSV_MIME = 'text/csv'

function stripExt(name: string): string {
  return name.replace(/\.(xlsx|xlsm|xltx|csv|xls)$/i, '')
}
function extOf(name: string): string {
  const m = /\.([a-z0-9]+)$/i.exec(name)
  return m ? m[1].toLowerCase() : 'xlsx'
}

// ---------- recent files (localStorage; per-viewer, best-effort) ----------
export interface RecentFile { name: string; path: string | null; date: number }
const RECENT_KEY = 'xl-recent-files'

export function getRecentFiles(): RecentFile[] {
  try {
    const raw = localStorage.getItem(RECENT_KEY)
    return raw ? (JSON.parse(raw) as RecentFile[]) : []
  } catch {
    return []
  }
}

function addRecentFile(name: string, path: string | null): void {
  try {
    const list = getRecentFiles().filter((f) => f.path !== path || !path)
    list.unshift({ name, path, date: Date.now() })
    localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, 25)))
  } catch {
    // localStorage unavailable (private mode, etc.) — recent list is just empty.
  }
}

function setTitle(): void {
  const title = `${store.fileName} - Excel`
  if (window.api) window.api.setTitle(title)
  else document.title = title
}

/** Shows Excel's "Want to save your changes?" prompt if needed. Returns false if the caller should abort. */
async function confirmDiscard(): Promise<boolean> {
  if (!store.dirty) return true
  const choice = await confirmSaveChanges(store.fileName)
  if (choice === 'cancel') return false
  if (choice === 'save') return save()
  return true
}

// ---------- browser fallbacks ----------
function browserPickFile(): Promise<{ name: string; data: Uint8Array } | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.xlsx,.xlsm,.csv'
    input.onchange = async () => {
      const f = input.files?.[0]
      if (!f) return resolve(null)
      const buf = new Uint8Array(await f.arrayBuffer())
      resolve({ name: f.name, data: buf })
    }
    // If the user cancels the native picker there is no reliable event; resolve(null) is the caller's fallback.
    input.click()
  })
}

function browserDownload(name: string, data: Uint8Array | string, mime: string): void {
  const blob = new Blob([data as BlobPart], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

// ---------- core operations ----------

export async function openFile(): Promise<void> {
  if (!(await confirmDiscard())) return

  if (window.api) {
    const r = await window.api.openFile()
    if (!r) return
    await loadFromBytes(r.data, r.path.split(/[\\/]/).pop() || 'Book1.xlsx', r.path)
  } else {
    const picked = await browserPickFile()
    if (!picked) return
    await loadFromBytes(picked.data, picked.name, null)
  }
}

async function loadFromBytes(data: Uint8Array, fileName: string, path: string | null): Promise<void> {
  const ext = extOf(fileName)
  const wb = ext === 'csv' ? readCsv(new TextDecoder('utf-8').decode(data)) : await readXlsx(data)
  store.load(wb, stripExt(fileName), path)
  setTitle()
  addRecentFile(fileName, path)
}

export async function save(): Promise<boolean> {
  if (!store.filePath && !window.api) return saveAs()
  if (!store.filePath) return saveAs()
  const bytes = await writeXlsx(store.toData())
  const ok = await window.api!.writeFile(store.filePath, bytes)
  if (ok) {
    store.dirty = false
    store.notify()
  }
  return ok
}

export async function saveAs(): Promise<boolean> {
  const currentName = store.fileName + '.xlsx'
  if (window.api) {
    const path = await window.api.saveAs(currentName)
    if (!path) return false
    const isCsv = extOf(path) === 'csv'
    const bytes = isCsv ? new TextEncoder().encode(writeActiveCsv()) : await writeXlsx(store.toData())
    const ok = await window.api.writeFile(path, bytes)
    if (ok) {
      store.fileName = stripExt(path.split(/[\\/]/).pop() || currentName)
      store.filePath = path
      store.dirty = false
      store.notify()
      setTitle()
      addRecentFile(path.split(/[\\/]/).pop() || currentName, path)
    }
    return ok
  }
  // Browser: no save dialog, just download.
  const bytes = await writeXlsx(store.toData())
  browserDownload(currentName, bytes, XLSX_MIME)
  store.dirty = false
  store.notify()
  return true
}

/** Exports the active sheet's displayed values as CSV text (Excel writes what's shown, not raw formulas). */
export function writeActiveCsv(): string {
  const sheet = store.active
  const used = store.usedRange(sheet)
  const rows = used ? used.r2 + 1 : 0
  const cols = used ? used.c2 + 1 : 0
  return writeCsv(rows, cols, (r, c) => store.getDisplay(sheet, r, c).text)
}

export function downloadCsv(): void {
  browserDownload(store.fileName + '.csv', writeActiveCsv(), CSV_MIME)
}

export async function newWorkbook(): Promise<void> {
  if (!(await confirmDiscard())) return
  store.newWorkbook()
  setTitle()
}

export async function closeWorkbook(): Promise<void> {
  if (!(await confirmDiscard())) return
  store.newWorkbook()
  setTitle()
}

// ---------- shortcuts & window-close guard ----------
let installed = false

export function installFileShortcuts(): void {
  if (installed) return
  installed = true

  window.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) && e.key !== 'F12') return
    const k = e.key.toLowerCase()
    if ((e.ctrlKey || e.metaKey) && k === 'n') { e.preventDefault(); void newWorkbook() }
    else if ((e.ctrlKey || e.metaKey) && k === 'o') { e.preventDefault(); void openFile() }
    else if ((e.ctrlKey || e.metaKey) && k === 's') { e.preventDefault(); void save() }
    else if (e.key === 'F12') { e.preventDefault(); void saveAs() }
    else if ((e.ctrlKey || e.metaKey) && k === 'w') { e.preventDefault(); void closeWorkbook() }
  })

  // Electron: the main process asks before actually closing the window when there are unsaved changes.
  window.api?.onCloseRequested?.(async () => {
    const ok = await confirmDiscard()
    await window.api!.confirmClose(ok)
  })
}

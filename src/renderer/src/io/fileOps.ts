// File menu operations: New, Open, Save, Save As, Close — Electron-backed, with a browser fallback.
import { store } from '../model/store'
import { readXlsx, writeXlsx } from './xlsx'
import { readCsv, writeCsv } from './csv'
import { confirmSaveChanges } from './saveConfirm'
import type { VbaProject } from './xlsxPostProcess'

// The currently-open workbook's VBA project (.xlsm only), kept outside WorkbookData/the store
// since it's opaque binary Excel data, not something the app model or undo stack needs to know
// about - just bytes to carry from open to the next save. Reset whenever a different file loads.
let currentVba: VbaProject | undefined

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

export function removeRecentFile(path: string): void {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(getRecentFiles().filter((f) => f.path !== path)))
  } catch {
    // localStorage unavailable: nothing to remove.
  }
}

function setTitle(): void {
  const title = `${store.fileName} - Excel`
  if (window.api) window.api.setTitle(title)
  else document.title = title
}

/** Tells main this window's current file name/path/dirty flag (used to decide window reuse on Open). */
function reportState(): void {
  window.api?.reportState?.({ fileName: store.fileName, filePath: store.filePath, dirty: store.dirty })
}

/** Shows Excel's "Want to save your changes?" prompt if needed. Returns false if the caller should abort. */
async function confirmDiscard(): Promise<boolean> {
  if (!store.dirty) return true
  const choice = await confirmSaveChanges(store.fileName)
  if (choice === 'cancel') return false
  if (choice === 'save') return save()
  clearRecovery() // "Don't Save": the user chose to drop these changes
  return true
}

// ---------- AutoRecover & AutoSave ----------
const recoverId = 'w' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8)
let recoverTimer: ReturnType<typeof setInterval> | null = null
let lastBackupVersion = -1

function clearRecovery(): void {
  lastBackupVersion = -1
  void window.api?.recoverClear?.(recoverId)
}

async function backupNow(): Promise<void> {
  if (!window.api?.recoverWrite || !store.dirty || store.getVersion() === lastBackupVersion) return
  lastBackupVersion = store.getVersion()
  try {
    const bytes = await writeXlsx(store.toData(), currentVba)
    await window.api.recoverWrite(recoverId, store.fileName, bytes)
  } catch { /* best effort */ }
}

let autoSaveOn = false
let autoSaveTimer: ReturnType<typeof setTimeout> | null = null
const autoSaveListeners = new Set<() => void>()
export function isAutoSaveOn(): boolean { return autoSaveOn }
export function onAutoSaveChange(fn: () => void): () => void { autoSaveListeners.add(fn); return () => autoSaveListeners.delete(fn) }
/** AutoSave: saves the open file a few seconds after each change. A workbook that has never been
 *  saved asks for a location first; CSV files are not auto-saved (a save drops everything but values). */
export async function setAutoSave(on: boolean): Promise<void> {
  if (on && (!store.filePath || extOf(store.filePath) === 'csv')) {
    if (!(await saveAs()) || !store.filePath || extOf(store.filePath) === 'csv') return
  }
  autoSaveOn = on
  autoSaveListeners.forEach((f) => f())
  if (on && store.dirty) void save()
}

function installRecovery(): void {
  void window.api?.recoverRegister?.(recoverId)
  recoverTimer ??= setInterval(() => void backupNow(), 60_000)
  store.subscribe(() => {
    if (!autoSaveOn || !store.dirty || !store.filePath) return
    if (autoSaveTimer) clearTimeout(autoSaveTimer)
    autoSaveTimer = setTimeout(() => { autoSaveTimer = null; if (autoSaveOn && store.dirty) void save() }, 3000)
  })
  // Reopen a backup left by a crash in the first window the user actually sees.
  const tryRecover = async (): Promise<void> => {
    if (document.visibilityState !== 'visible' || store.dirty || store.filePath) return
    document.removeEventListener('visibilitychange', tryRecover)
    const rec = await window.api?.recoverTakeOne?.()
    if (!rec) return
    try {
      await loadFromBytes(rec.data, rec.name + ' (Recovered).xlsx', null)
      store.fileName = rec.name + ' (Recovered)'
      store.dirty = true
      store.notify()
      setTitle()
      window.alert(`Cellbook closed unexpectedly last time. Your unsaved work in '${rec.name}' was recovered.\n\nSave it now with File > Save As.`)
    } catch { /* unreadable backup */ }
  }
  document.addEventListener('visibilitychange', tryRecover)
  void tryRecover()
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

/**
 * Returns true when the pick actually went somewhere (Electron: main routes it to a window —
 * this one or a new one — asynchronously; browser: loads into this tab synchronously-ish).
 */
export async function openFile(): Promise<boolean> {
  if (window.api) return window.api.pickAndOpen()

  if (!(await confirmDiscard())) return false
  const picked = await browserPickFile()
  if (!picked) return false
  try {
    await loadFromBytes(picked.data, picked.name, null)
  } catch {
    window.alert(`Excel cannot open the file '${picked.name}' because the file format or file extension is not valid. Verify that the file has not been corrupted and that the file extension matches the format of the file.`)
    return false
  }
  return true
}

/** Opens a file from the Recent list. In the browser (no paths) it falls back to the picker; a
 *  file that has been moved or deleted gets Excel's "couldn't find" prompt and can be removed. */
export async function openRecent(f: RecentFile): Promise<boolean> {
  if (!window.api || !f.path) return openFile()
  if (await window.api.openPathRequest(f.path)) return true
  if (window.confirm(`Sorry, we couldn't find ${f.path}. Is it possible it was moved, renamed or deleted?\n\nRemove it from the Recent list?`)) removeRecentFile(f.path)
  return false
}

async function loadFromBytes(data: Uint8Array, fileName: string, path: string | null): Promise<void> {
  const ext = extOf(fileName)
  if (ext === 'csv') {
    currentVba = undefined
    store.load(readCsv(new TextDecoder('utf-8').decode(data)), stripExt(fileName), path)
  } else {
    const wb = await readXlsx(data)
    currentVba = wb.vbaProject
    store.load(wb, stripExt(fileName), path)
  }
  setTitle()
  addRecentFile(fileName, path)
}

/** Only a .xlsm save carries the macros back; Save As to .xlsx must drop them, same as real Excel. */
function vbaForPath(path: string): VbaProject | undefined {
  return extOf(path) === 'xlsm' ? currentVba : undefined
}

/** Writes bytes to disk; on failure tells the user (Excel's wording) and returns false. */
async function writeOrReport(path: string, bytes: Uint8Array): Promise<boolean> {
  let res: true | string
  try {
    res = await window.api!.writeFile(path, bytes)
  } catch (e) {
    res = (e as Error)?.message || 'Unknown error'
  }
  if (res === true) return true
  const name = path.split(/[\\/]/).pop()
  window.alert(`Your changes couldn't be saved to '${name}'.\n\n${res}\n\nYour work is still open. Try Save As to save a copy in another location.`)
  return false
}

let saving: Promise<boolean> | null = null

export async function save(): Promise<boolean> {
  if (!store.filePath) return saveAs()
  if (saving) return saving // Ctrl+S held down / AutoSave overlapping: one save at a time
  const path = store.filePath
  const version = store.getVersion()
  saving = (async () => {
    const bytes = extOf(path) === 'csv'
      ? new TextEncoder().encode(writeActiveCsv())
      : await writeXlsx(store.toData(), vbaForPath(path))
    const ok = await writeOrReport(path, bytes)
    // Only edits captured in this save count as saved; anything typed meanwhile stays unsaved.
    if (ok && store.getVersion() === version && store.filePath === path) {
      store.dirty = false
      store.notify()
      window.api?.reportWorkbookSaved?.({ fileName: store.fileName, filePath: store.filePath })
      clearRecovery()
    }
    return ok
  })()
  try { return await saving } finally { saving = null }
}

export async function saveAs(): Promise<boolean> {
  const currentName = store.fileName + '.xlsx'
  if (window.api) {
    const path = await window.api.saveAs(currentName)
    if (!path) return false
    const isCsv = extOf(path) === 'csv'
    const bytes = isCsv ? new TextEncoder().encode(writeActiveCsv()) : await writeXlsx(store.toData(), vbaForPath(path))
    const ok = await writeOrReport(path, bytes)
    if (ok) {
      clearRecovery()
      store.fileName = stripExt(path.split(/[\\/]/).pop() || currentName)
      store.filePath = path
      store.dirty = false
      store.notify()
      setTitle()
      addRecentFile(path.split(/[\\/]/).pop() || currentName, path)
      window.api?.reportWorkbookSaved?.({ fileName: store.fileName, filePath: store.filePath })
    }
    return ok
  }
  // Browser: no save dialog, just download (always .xlsx, so no macros to carry over).
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
  const text = writeCsv(rows, cols, (r, c) => store.getDisplay(sheet, r, c).text)
  // Non-ASCII text (Hindi, ₹, ...) needs the UTF-8 BOM, as in Excel's "CSV UTF-8", or Excel shows it garbled.
  return /[^\x00-\x7f]/.test(text) ? '\uFEFF' + text : text
}

export function downloadCsv(): void {
  browserDownload(store.fileName + '.csv', writeActiveCsv(), CSV_MIME)
}

/**
 * Ctrl+N / File > New > Blank workbook: Excel always opens a NEW window for this, it never
 * reuses the current one. In the browser fallback (no window.api) there's only one "window",
 * so it resets in place instead.
 */
export async function newWorkbook(): Promise<boolean> {
  if (window.api) {
    await window.api.newWindow()
    return true
  }
  if (!(await confirmDiscard())) return false
  currentVba = undefined
  store.newWorkbook()
  setTitle()
  return true
}

/** Ctrl+W / Ctrl+F4: closes this workbook's window (with the save prompt); the last window quits the app. */
export async function closeWorkbook(): Promise<boolean> {
  if (window.api) {
    await window.api.windowControl('close') // main re-asks via onCloseRequested below if there are unsaved changes
    return true
  }
  if (!(await confirmDiscard())) return false
  currentVba = undefined
  store.newWorkbook()
  setTitle()
  return true
}

// ---------- shortcuts & window-close guard ----------
let installed = false

export function installFileShortcuts(): void {
  if (installed) return
  installed = true
  installRecovery()

  window.addEventListener('keydown', (e) => {
    const ctrl = e.ctrlKey || e.metaKey
    if (!ctrl && e.key !== 'F12') return
    const k = e.key.toLowerCase()
    if (ctrl && k === 'n') { e.preventDefault(); void newWorkbook() }
    else if (ctrl && k === 'o') { e.preventDefault(); void openFile() }
    else if (ctrl && k === 's') { e.preventDefault(); void save() }
    else if (e.key === 'F12') { e.preventDefault(); void saveAs() }
    else if (ctrl && (k === 'w' || e.key === 'F4')) { e.preventDefault(); void closeWorkbook() }
    else if (ctrl && (k === 'tab' || e.key === 'F6')) { e.preventDefault(); void window.api?.cycleWindow(e.shiftKey ? -1 : 1) }
  })

  // Electron: the main process asks before actually closing the window when there are unsaved changes.
  window.api?.onCloseRequested?.(async () => {
    const ok = await confirmDiscard()
    await window.api!.confirmClose(ok)
  })

  // A spare pre-loaded window is handed its "BookN" name only once it's actually shown.
  window.api?.onAssignBook?.((name) => {
    store.fileName = name
    store.notify()
    setTitle()
    reportState()
  })

  // External opens: file associations, second-instance launches, Recent, and Ctrl+O all funnel
  // through main, which pushes the bytes here once it has decided which window they belong in.
  window.api?.onOpenPath?.(async ({ path, data }) => {
    if (!(await confirmDiscard())) return
    const name = path.split(/[\\/]/).pop() || 'Book1.xlsx'
    try {
      await loadFromBytes(data, name, path)
    } catch {
      // Corrupt file / wrong format, e.g. a renamed non-workbook file — leave this window's
      // current workbook untouched instead of crashing, matching Excel's own error prompt.
      window.alert(`Excel cannot open the file '${name}' because the file format or file extension is not valid. Verify that the file has not been corrupted and that the file extension matches the format of the file.`)
    }
  })

  // Multi-window view synchronization
  store.onBroadcast = (msg) => {
    window.api?.broadcastOp?.(msg)
  }

  window.api?.onRemoteOp?.((msg) => {
    store.applyRemoteSync(msg as any)
  })

  window.api?.onWorkbookSaved?.((info) => {
    store.fileName = info.fileName
    store.filePath = info.filePath
    store.dirty = false
    store.notify()
    setTitle()
  })

  window.api?.onInitView?.((payload) => {
    store.load(payload.data as any, payload.fileName, payload.filePath)
    setTitle()
  })

  // Report this window's state whenever the workbook changes (name/path/dirty), so main can
  // decide window reuse correctly and keep its Switch Windows list in sync.
  store.subscribe(reportState)
  reportState()

  installDropToOpen()
}

/** Drag a workbook file onto the window to open it, like dropping it onto Excel's icon. */
function installDropToOpen(): void {
  window.addEventListener('dragover', (e) => {
    if (e.dataTransfer?.types.includes('Files')) e.preventDefault()
  })
  window.addEventListener('drop', (e) => {
    const files = e.dataTransfer?.files
    if (!files?.length || !window.api?.getPathForFile) return
    e.preventDefault()
    for (const f of Array.from(files)) {
      const path = window.api.getPathForFile(f)
      if (/\.(xlsx|xlsm|xls|csv)$/i.test(path)) void window.api.openPathRequest(path)
    }
  })
}

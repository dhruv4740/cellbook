import { app, BrowserWindow, clipboard, ClipboardItem, dialog, ipcMain, nativeImage, screen, shell } from 'electron'
import { join, basename } from 'path'
import { promises as fs, existsSync } from 'fs'
import { canHideWindow, hiddenWindows, WinVis } from './windowVisibility'

// Chromium gives pages the Windows display language (often en-US), but Excel reads dates and
// digit grouping from the Windows region (e.g. English (India): 3/9/2026 = 3 Sept, 12,34,567).
// Node's default locale is that region; the preload hands it to the renderer (engine/locale.ts).
const regionArg = `--cellbook-region=${Intl.DateTimeFormat().resolvedOptions().locale}`

// ---------------------------------------------------------------------------
// One workbook per window, like Excel: each BrowserWindow hosts exactly one
// workbook (or a secondary view of a workbook via View > New Window).
// ---------------------------------------------------------------------------

interface WinState {
  fileName: string
  filePath: string | null
  dirty: boolean
}

interface WinEntry {
  win: BrowserWindow
  state: WinState
  allowClose: boolean
  workbookId: string
  viewIndex: number
}

const windows = new Map<number, WinEntry>()
/** Creation order, used for Ctrl+Tab / Ctrl+F6 cycling. */
const windowOrder: number[] = []
let bookCounter = 0
let wbCounter = 0

function nextWbId(): string {
  wbCounter += 1
  return `wb_${wbCounter}`
}

/** A hidden, pre-loaded window kept ready so a plain Ctrl+N feels instant. */
let spare: BrowserWindow | null = null
/** Spare windows whose renderer has finished loading (and so is listening for assignBook). */
const loadedSpares = new WeakSet<BrowserWindow>()

function preloadSpareWindow(): void {
  const win = new BrowserWindow({
    show: false,
    width: 1400,
    height: 900,
    minWidth: 800,
    minHeight: 500,
    frame: false,
    backgroundColor: '#ffffff',
    webPreferences: { preload: join(__dirname, '../preload/index.js'), sandbox: false, additionalArguments: [regionArg] }
  })
  win.webContents.once('did-finish-load', () => loadedSpares.add(win))
  if (process.env.ELECTRON_RENDERER_URL) win.loadURL(process.env.ELECTRON_RENDERER_URL)
  else win.loadFile(join(__dirname, '../renderer/index.html'))
  spare = win
}

const filters = [
  { name: 'Excel Workbook', extensions: ['xlsx'] },
  { name: 'Excel Macro-Enabled Workbook', extensions: ['xlsm'] },
  { name: 'CSV (Comma delimited)', extensions: ['csv'] },
  { name: 'All Files', extensions: ['*'] }
]

function nextBookName(): string {
  bookCounter += 1
  return `Book${bookCounter}`
}

function isUntouchedBook(s: WinState): boolean {
  return s.filePath === null && !s.dirty && /^Book\d+$/.test(s.fileName)
}

function samePath(a: string, b: string): boolean {
  return process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b
}

function findWindowForPath(filePath: string): WinEntry | undefined {
  for (const entry of windows.values()) {
    if (entry.state.filePath && samePath(entry.state.filePath, filePath)) return entry
  }
  return undefined
}

function broadcastWindowList(): void {
  const list = getWindowList()
  for (const entry of windows.values()) entry.win.webContents.send('window:listChanged', list)
}

function getWindowList(): { id: number; title: string; current: boolean }[] {
  const focused = BrowserWindow.getFocusedWindow()
  return windowOrder
    .filter((id) => windows.has(id))
    .map((id) => {
      const entry = windows.get(id)!
      const suffix = entry.viewIndex > 0 ? `:${entry.viewIndex}` : ''
      return { id, title: `${entry.state.fileName}${suffix} - Excel`, current: focused?.id === id }
    })
}

function updateWorkbookViewTitles(wbId: string): void {
  const entries = windowOrder.filter((id) => windows.has(id) && windows.get(id)!.workbookId === wbId).map((id) => windows.get(id)!)
  if (entries.length <= 1) {
    for (const e of entries) {
      e.viewIndex = 0
      const title = `${e.state.fileName} - Excel`
      e.win.setTitle(title)
      e.win.webContents.send('window:viewIndexChanged', 0)
    }
  } else {
    entries.forEach((e, idx) => {
      const viewNum = idx + 1
      e.viewIndex = viewNum
      const title = `${e.state.fileName}:${viewNum} - Excel`
      e.win.setTitle(title)
      e.win.webContents.send('window:viewIndexChanged', viewNum)
    })
  }
  broadcastWindowList()
}

/** Creates a new workbook window. Exactly one of `book`/`filePath` should be given. */
function createWindow(opts: { book?: string; filePath?: string; workbookId?: string } = {}): BrowserWindow {
  const bookName = opts.book ?? (opts.filePath ? basename(opts.filePath) : nextBookName())
  const workbookId = opts.workbookId ?? nextWbId()
  let win: BrowserWindow
  let usedSpare = false

  if (!opts.filePath && !opts.workbookId && spare) {
    win = spare
    spare = null
    usedSpare = true
    win.setTitle(`${bookName} - Excel`)
    // A very fast Ctrl+N can reuse the spare before its renderer has loaded; hold the name
    // (and the reveal) until it is listening, or it would keep its default Book1 name.
    const assign = (w: BrowserWindow): void => { w.webContents.send('window:assignBook', bookName); w.show() }
    if (loadedSpares.has(win)) assign(win)
    else win.webContents.once('did-finish-load', () => assign(win))
    // Keep a spare ready for the *next* Ctrl+N.
    preloadSpareWindow()
  } else {
    win = new BrowserWindow({
      width: 1400,
      height: 900,
      minWidth: 800,
      minHeight: 500,
      frame: false, // Excel draws its own title bar
      backgroundColor: '#ffffff',
      title: `${bookName} - Excel`,
      webPreferences: { preload: join(__dirname, '../preload/index.js'), sandbox: false, additionalArguments: [regionArg] }
    })
    const query: Record<string, string> = opts.filePath ? {} : { book: bookName }
    if (process.env.ELECTRON_RENDERER_URL) {
      const url = new URL(process.env.ELECTRON_RENDERER_URL)
      for (const [k, v] of Object.entries(query)) url.searchParams.set(k, v)
      win.loadURL(url.toString())
    } else {
      win.loadFile(join(__dirname, '../renderer/index.html'), { query })
    }
  }

  const entry: WinEntry = {
    win,
    state: { fileName: bookName, filePath: opts.filePath ?? null, dirty: false },
    allowClose: false,
    workbookId,
    viewIndex: 0
  }
  windows.set(win.id, entry)
  windowOrder.push(win.id)

  win.on('maximize', () => win.webContents.send('window:maximized', true))
  win.on('unmaximize', () => win.webContents.send('window:maximized', false))
  win.on('focus', () => broadcastWindowList())

  // Give the renderer a chance to prompt "Save changes?" before the window actually closes.
  win.on('close', (e) => {
    if (entry.allowClose) return
    const siblings = [...windows.values()].filter((w) => w.workbookId === entry.workbookId && w.win.id !== entry.win.id)
    if (siblings.length > 0) {
      // Not the last window for this workbook: allow immediate close without save prompt
      entry.allowClose = true
      return
    }
    if (!entry.state.dirty) {
      entry.allowClose = true
      return
    }
    // Last window and unsaved changes: prompt
    e.preventDefault()
    win.webContents.send('window:closeRequested')
  })
  win.on('closed', () => {
    const wbId = entry.workbookId
    windows.delete(win.id)
    const idx = windowOrder.indexOf(win.id)
    if (idx >= 0) windowOrder.splice(idx, 1)
    updateWorkbookViewTitles(wbId)
    // The spare hidden window doesn't count as a real workbook window, so
    // 'window-all-closed' won't fire on its own once the last real one closes.
    if (windows.size === 0) {
      spare?.destroy()
      spare = null
      app.quit()
    }
  })

  if (opts.filePath && !usedSpare) {
    win.webContents.once('did-finish-load', () => {
      void openPathInWindow(entry, opts.filePath!)
    })
  }

  updateWorkbookViewTitles(workbookId)
  return win
}

/** Reads `filePath` and pushes it into an already-created window's renderer. */
async function openPathInWindow(entry: WinEntry, filePath: string): Promise<void> {
  try {
    const data = await fs.readFile(filePath)
    entry.state = { fileName: basename(filePath), filePath, dirty: false }
    entry.win.webContents.send('file:openPath', { path: filePath, data: new Uint8Array(data) })
    updateWorkbookViewTitles(entry.workbookId)
  } catch {
    dialog.showErrorBox('Cellbook', `Couldn't open "${filePath}".`)
  }
}

/**
 * Central "open this file somewhere" router used by Ctrl+O, Recent, file
 * associations, drag-drop and second-instance launches:
 *  - if already open, focus that window;
 *  - else reuse `originWin` (or the sole open window) if it's an untouched Book1;
 *  - else open a new window.
 */
async function openPathSomewhere(filePath: string, originWin?: BrowserWindow): Promise<void> {
  if (!existsSync(filePath)) {
    dialog.showErrorBox('Cellbook', `Couldn't find "${filePath}".`)
    return
  }
  const already = findWindowForPath(filePath)
  if (already) {
    if (already.win.isMinimized()) already.win.restore()
    already.win.show()
    already.win.focus()
    return
  }

  let target: WinEntry | undefined
  if (originWin) {
    const e = windows.get(originWin.id)
    if (e && isUntouchedBook(e.state)) target = e
  }
  if (!target && windows.size === 1) {
    const only = [...windows.values()][0]
    if (isUntouchedBook(only.state)) target = only
  }

  if (target) {
    await openPathInWindow(target, filePath)
    target.win.show()
    target.win.focus()
  } else {
    createWindow({ filePath })
  }
}

// ---------------------------------------------------------------------------
// IPC: per-window controls
// ---------------------------------------------------------------------------

ipcMain.handle('window:control', (e, action: 'minimize' | 'maximize' | 'close') => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (!win) return
  if (action === 'minimize') win.minimize()
  else if (action === 'maximize') (win.isMaximized() ? win.unmaximize() : win.maximize())
  else win.close()
})
ipcMain.handle('window:setTitle', (e, title: string) => BrowserWindow.fromWebContents(e.sender)?.setTitle(title))

ipcMain.handle('window:confirmClose', (e, shouldClose: boolean) => {
  const win = BrowserWindow.fromWebContents(e.sender)
  const entry = win && windows.get(win.id)
  if (!entry) return
  if (shouldClose) {
    entry.allowClose = true
    entry.win.close()
  }
})

ipcMain.on('window:reportState', (e, state: WinState) => {
  const win = BrowserWindow.fromWebContents(e.sender)
  const entry = win && windows.get(win.id)
  if (!entry) return
  const changed = entry.state.fileName !== state.fileName || entry.state.filePath !== state.filePath || entry.state.dirty !== state.dirty
  entry.state = state
  if (!changed) return
  updateWorkbookViewTitles(entry.workbookId)
})

// ---------------------------------------------------------------------------
// IPC: multi-window management (View > Window group, Ctrl+Tab, Ctrl+N)
// ---------------------------------------------------------------------------

ipcMain.handle('window:new', (_e, _opts?: { isNewWorkbook?: boolean }) => {
  createWindow()
})

ipcMain.handle('window:newView', (e, snapshot: { data: unknown; fileName: string; filePath: string | null }) => {
  const senderWin = BrowserWindow.fromWebContents(e.sender)
  const senderEntry = senderWin && windows.get(senderWin.id)
  if (!senderWin || !senderEntry) return
  const wbId = senderEntry.workbookId
  const bookName = senderEntry.state.fileName

  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 800,
    minHeight: 500,
    frame: false,
    backgroundColor: '#ffffff',
    title: `${bookName} - Excel`,
    webPreferences: { preload: join(__dirname, '../preload/index.js'), sandbox: false, additionalArguments: [regionArg] }
  })

  const entry: WinEntry = {
    win,
    state: { fileName: bookName, filePath: senderEntry.state.filePath, dirty: senderEntry.state.dirty },
    allowClose: false,
    workbookId: wbId,
    viewIndex: 0
  }
  windows.set(win.id, entry)
  windowOrder.push(win.id)

  win.on('maximize', () => win.webContents.send('window:maximized', true))
  win.on('unmaximize', () => win.webContents.send('window:maximized', false))
  win.on('focus', () => broadcastWindowList())

  win.on('close', (evt) => {
    if (entry.allowClose) return
    const siblings = [...windows.values()].filter((w) => w.workbookId === entry.workbookId && w.win.id !== entry.win.id)
    if (siblings.length > 0) {
      entry.allowClose = true
      return
    }
    if (!entry.state.dirty) {
      entry.allowClose = true
      return
    }
    evt.preventDefault()
    win.webContents.send('window:closeRequested')
  })

  win.on('closed', () => {
    windows.delete(win.id)
    const idx = windowOrder.indexOf(win.id)
    if (idx >= 0) windowOrder.splice(idx, 1)
    updateWorkbookViewTitles(wbId)
    if (windows.size === 0) {
      spare?.destroy()
      spare = null
      app.quit()
    }
  })

  win.webContents.once('did-finish-load', () => {
    updateWorkbookViewTitles(wbId)
    win.webContents.send('window:initView', {
      data: snapshot.data,
      fileName: snapshot.fileName,
      filePath: snapshot.filePath,
      viewIndex: entry.viewIndex
    })
    win.show()
  })

  const query: Record<string, string> = { book: bookName }
  if (process.env.ELECTRON_RENDERER_URL) {
    const url = new URL(process.env.ELECTRON_RENDERER_URL)
    for (const [k, v] of Object.entries(query)) url.searchParams.set(k, v)
    win.loadURL(url.toString())
  } else {
    win.loadFile(join(__dirname, '../renderer/index.html'), { query })
  }
})

ipcMain.on('workbook:op', (e, msg) => {
  const senderWin = BrowserWindow.fromWebContents(e.sender)
  const senderEntry = senderWin && windows.get(senderWin.id)
  if (!senderEntry) return
  for (const entry of windows.values()) {
    if (entry.win.webContents.id !== e.sender.id && entry.workbookId === senderEntry.workbookId) {
      entry.win.webContents.send('workbook:op', msg)
    }
  }
})

ipcMain.on('workbook:saved', (e, info: { fileName: string; filePath: string | null }) => {
  const senderWin = BrowserWindow.fromWebContents(e.sender)
  const senderEntry = senderWin && windows.get(senderWin.id)
  if (!senderEntry) return
  for (const entry of windows.values()) {
    if (entry.workbookId === senderEntry.workbookId) {
      entry.state.fileName = info.fileName
      entry.state.filePath = info.filePath
      entry.state.dirty = false
      if (entry.win.webContents.id !== e.sender.id) {
        entry.win.webContents.send('workbook:saved', info)
      }
    }
  }
  updateWorkbookViewTitles(senderEntry.workbookId)
})

ipcMain.handle('window:list', () => getWindowList())

const winVis = (): WinVis[] => getWindowList().map((w) => ({ id: w.id, title: w.title, visible: windows.get(w.id)!.win.isVisible() }))

/** View > Hide: hides the sender's window. Returns false (and does nothing) if it is the only visible one. */
ipcMain.handle('window:hide', (e): boolean => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (!win || !canHideWindow(winVis(), win.id)) return false
  win.hide()
  return true
})
ipcMain.handle('window:listHidden', () => hiddenWindows(winVis()))

ipcMain.handle('window:activate', (_e, id: number) => {
  const entry = windows.get(id)
  if (!entry) return
  if (entry.win.isMinimized()) entry.win.restore()
  entry.win.show()
  entry.win.focus()
})

ipcMain.handle('window:cycle', (e, dir: 1 | -1) => {
  const cur = BrowserWindow.fromWebContents(e.sender)
  const ids = windowOrder.filter((id) => windows.has(id))
  if (ids.length < 2 || !cur) return
  const i = ids.indexOf(cur.id)
  const next = ids[(i + dir + ids.length) % ids.length]
  const entry = windows.get(next)
  if (entry) {
    if (entry.win.isMinimized()) entry.win.restore()
    entry.win.show()
    entry.win.focus()
  }
})

ipcMain.handle('window:arrangePrompt', async (e) => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (!win) return
  const r = await dialog.showMessageBox(win, {
    type: 'question',
    title: 'Arrange Windows',
    message: 'Arrange Windows',
    buttons: ['Tiled', 'Horizontal', 'Vertical', 'Cascade', 'Cancel'],
    defaultId: 0,
    cancelId: 4,
    noLink: true
  })
  const modes = ['tiled', 'horizontal', 'vertical', 'cascade'] as const
  if (r.response < 4) arrangeWindows(modes[r.response])
})

ipcMain.handle('window:arrange', (_e, mode: 'tiled' | 'horizontal' | 'vertical' | 'cascade') => {
  arrangeWindows(mode)
})

function arrangeWindows(mode: 'tiled' | 'horizontal' | 'vertical' | 'cascade'): void {
  const entries = windowOrder.filter((id) => windows.has(id)).map((id) => windows.get(id)!)
  if (!entries.length) return
  const area = screen.getPrimaryDisplay().workArea
  const n = entries.length

  entries.forEach((entry, i) => {
    if (entry.win.isMinimized()) entry.win.restore()
    if (mode === 'cascade') {
      const off = i * 32
      const w = Math.round(area.width * 0.6)
      const h = Math.round(area.height * 0.6)
      entry.win.setBounds({ x: area.x + off, y: area.y + off, width: w, height: h })
    } else if (mode === 'horizontal') {
      const h = Math.floor(area.height / n)
      entry.win.setBounds({ x: area.x, y: area.y + i * h, width: area.width, height: h })
    } else if (mode === 'vertical') {
      const w = Math.floor(area.width / n)
      entry.win.setBounds({ x: area.x + i * w, y: area.y, width: w, height: area.height })
    } else {
      // Tiled: a roughly-square grid, Excel-style.
      const cols = Math.ceil(Math.sqrt(n))
      const rows = Math.ceil(n / cols)
      const w = Math.floor(area.width / cols)
      const h = Math.floor(area.height / rows)
      const col = i % cols
      const row = Math.floor(i / cols)
      entry.win.setBounds({ x: area.x + col * w, y: area.y + row * h, width: w, height: h })
    }
    entry.win.show()
  })
}

// ---------------------------------------------------------------------------
// IPC: file open/save
// ---------------------------------------------------------------------------

/** Ctrl+O / File > Open > Browse: shows the dialog, then routes to the right window. */
ipcMain.handle('file:pickAndOpen', async (e) => {
  const win = BrowserWindow.fromWebContents(e.sender)!
  const r = await dialog.showOpenDialog(win, { properties: ['openFile', 'multiSelections'], filters })
  if (r.canceled || !r.filePaths.length) return false
  for (const p of r.filePaths) await openPathSomewhere(p, win)
  return true
})

/** Recent list / drag-drop with a known path already in hand. */
ipcMain.handle('file:openPathRequest', async (e, filePath: string) => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (!existsSync(filePath)) return false // the renderer explains and offers to drop it from Recent
  await openPathSomewhere(filePath, win ?? undefined)
  return true
})

ipcMain.handle('file:saveAs', async (e, defaultName: string) => {
  const win = BrowserWindow.fromWebContents(e.sender)!
  const r = await dialog.showSaveDialog(win, { defaultPath: defaultName, filters })
  return r.canceled ? null : r.filePath
})

// System clipboard through main: navigator.clipboard in the renderer can stall on permission checks.
ipcMain.handle('spell:addWord', (e, word: string) => e.sender.session.addWordToSpellCheckerDictionary(word))
ipcMain.handle('spell:languages', (e) => e.sender.session.getSpellCheckerLanguages())
ipcMain.handle('clipboard:readText', () => clipboard.readText())
ipcMain.handle('clipboard:readImage', async () => {
  try {
    for (const item of await clipboard.read()) {
      if (item.types.includes('image/png')) return 'data:image/png;base64,' + Buffer.from(await ((await item.getType('image/png')) as Blob).arrayBuffer()).toString('base64')
    }
  } catch {
    // no image on the clipboard
  }
  return ''
})
ipcMain.handle('clipboard:readHTML', async () => {
  try {
    for (const item of await clipboard.read()) {
      if (item.types.includes('text/html')) return await ((await item.getType('text/html')) as Blob).text()
    }
  } catch {
    // no HTML on the clipboard
  }
  return ''
})
ipcMain.handle('clipboard:write', (_e, data: { text: string; html?: string }) =>
  clipboard.write([new ClipboardItem(data.html ? { 'text/plain': data.text, 'text/html': data.html } : { 'text/plain': data.text })])
)
ipcMain.handle('clipboard:writeImage', (_e, dataUrl: string) => clipboard.write([new ClipboardItem({ 'image/png': new Blob([new Uint8Array(nativeImage.createFromDataURL(dataUrl).toPNG())], { type: 'image/png' }) })]))
ipcMain.handle('file:read', async (_e, path: string) => new Uint8Array(await fs.readFile(path)))

// AutoRecover: while a workbook has unsaved changes its window writes a backup here every minute;
// a clean save or close removes it. Backups left behind by a crash are reopened on the next launch.
const recoverDir = (): string => join(app.getPath('userData'), 'AutoRecover')
const safeId = (id: string): string => id.replace(/[^A-Za-z0-9_-]/g, '')
ipcMain.handle('recover:write', async (_e, id: string, name: string, data: Uint8Array) => {
  try {
    await fs.mkdir(recoverDir(), { recursive: true })
    const base = join(recoverDir(), safeId(id))
    await fs.writeFile(base + '.tmp', data)
    await fs.rename(base + '.tmp', base + '.xlsx')
    await fs.writeFile(base + '.json', JSON.stringify({ name, time: Date.now() }))
  } catch { /* best effort */ }
})
ipcMain.handle('recover:clear', async (_e, id: string) => {
  const base = join(recoverDir(), safeId(id))
  await Promise.all(['.xlsx', '.json', '.tmp'].map((x) => fs.rm(base + x, { force: true }).catch(() => {})))
})
/** Hands one orphaned backup (from a previous run) to the asking window and removes it from disk. */
const liveRecoverIds = new Set<string>()
ipcMain.handle('recover:register', (_e, id: string) => { liveRecoverIds.add(safeId(id)) })
ipcMain.handle('recover:takeOne', async () => {
  try {
    const files = (await fs.readdir(recoverDir())).filter((f) => f.endsWith('.xlsx'))
    for (const f of files) {
      const id = f.slice(0, -5)
      if (liveRecoverIds.has(id)) continue
      const base = join(recoverDir(), id)
      const data = new Uint8Array(await fs.readFile(base + '.xlsx'))
      let name = 'Book1'
      try { name = JSON.parse(await fs.readFile(base + '.json', 'utf8')).name || name } catch { /* no meta */ }
      await Promise.all(['.xlsx', '.json'].map((x) => fs.rm(base + x, { force: true }).catch(() => {})))
      return { name, data }
    }
  } catch { /* no folder yet */ }
  return null
})

// Clicking a cell hyperlink: web and mail links open in the default app; anything else is treated
// as a file path (Excel opens those with their associated program). Returns an error, or true.
ipcMain.handle('shell:openLink', async (_e, target: string): Promise<true | string> => {
  try {
    if (/^(https?:|mailto:)/i.test(target)) { await shell.openExternal(target); return true }
    const p = target.replace(/^file:\/\/\/?/i, '')
    const err = await shell.openPath(decodeURIComponent(p))
    return err || true
  } catch (err) {
    return (err as Error).message || String(err)
  }
})

// File > Print / Export PDF: the renderer builds a print document of the sheet; main renders it in a
// hidden window and either opens the system print dialog or saves a PDF.
ipcMain.handle('print:run', async (e, opts: { html: string; mode: 'print' | 'pdf'; widthIn: number; heightIn: number; defaultName: string }): Promise<true | string | null> => {
  const owner = BrowserWindow.fromWebContents(e.sender) ?? undefined
  const tmp = join(app.getPath('temp'), `cellbook-print-${process.pid}-${Date.now()}.html`)
  const pw = new BrowserWindow({ show: false, webPreferences: { javascript: false, sandbox: true } })
  try {
    await fs.writeFile(tmp, opts.html, 'utf8')
    await pw.loadFile(tmp)
    if (opts.mode === 'pdf') {
      const target = owner
        ? await dialog.showSaveDialog(owner, { defaultPath: opts.defaultName.replace(/\.[^.]+$/, '') + '.pdf', filters: [{ name: 'PDF', extensions: ['pdf'] }] })
        : await dialog.showSaveDialog({ defaultPath: opts.defaultName + '.pdf' })
      if (target.canceled || !target.filePath) return null
      const pdf = await pw.webContents.printToPDF({ pageSize: { width: opts.widthIn, height: opts.heightIn }, printBackground: true, preferCSSPageSize: true })
      await fs.writeFile(target.filePath, pdf)
      return true
    }
    return await new Promise<true | string | null>((resolve) => {
      pw.webContents.print({ silent: false, printBackground: true, landscape: false, pageSize: { width: Math.round(opts.widthIn * 25400), height: Math.round(opts.heightIn * 25400) } }, (ok, reason) => {
        resolve(ok ? true : reason === 'cancelled' || reason === 'Print job canceled' ? null : reason || 'Printing failed.')
      })
    })
  } catch (err) {
    return (err as Error).message || String(err)
  } finally {
    pw.destroy()
    fs.rm(tmp, { force: true }).catch(() => {})
  }
})

// Safe save, like Excel: write a temporary file next to the target, then swap it in, so a failed
// or interrupted save never leaves a half-written workbook. Returns an error message on failure.
ipcMain.handle('file:write', async (_e, path: string, data: Uint8Array): Promise<true | string> => {
  const tmp = path + '.~cbtmp' + process.pid
  try {
    await fs.writeFile(tmp, data)
    await fs.rename(tmp, path)
    return true
  } catch (err) {
    await fs.rm(tmp, { force: true }).catch(() => {})
    const code = (err as NodeJS.ErrnoException).code
    if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES')
      return 'The file may be open in another program, read-only, or you may not have permission to save in this location.'
    if (code === 'ENOSPC') return 'There is not enough disk space.'
    if (code === 'ENOENT') return 'The folder no longer exists or the drive is not available.'
    return (err as Error).message || String(err)
  }
})

// ---------------------------------------------------------------------------
// Startup / single instance / file associations
// ---------------------------------------------------------------------------

/** Args that look like a workbook path (ignores flags and the electron/app entry itself). */
function fileArgsFrom(argv: string[]): string[] {
  return argv.filter((a) => !a.startsWith('-') && /\.(xlsx|xlsm|xls|csv)$/i.test(a))
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_e, argv) => {
    const paths = fileArgsFrom(argv)
    if (paths.length) {
      for (const p of paths) void openPathSomewhere(p)
    } else if (windows.size === 0) {
      createWindow()
    } else {
      // No file given: bring the most recently focused window forward, like Excel does.
      const last = windows.get(windowOrder[windowOrder.length - 1])
      last?.win.show()
      last?.win.focus()
    }
  })

  app.whenReady().then(() => {
    const paths = fileArgsFrom(process.argv.slice(app.isPackaged ? 1 : 2))
    if (paths.length) {
      for (const p of paths) void openPathSomewhere(p)
    } else {
      createWindow()
    }
    preloadSpareWindow()
  })
}

app.on('window-all-closed', () => app.quit())

import { app, BrowserWindow, dialog, ipcMain, screen } from 'electron'
import { join, basename } from 'path'
import { promises as fs, existsSync } from 'fs'

// ---------------------------------------------------------------------------
// One workbook per window, like Excel: each BrowserWindow hosts exactly one
// workbook. This module tracks every open window, the workbook state the
// renderer last reported (name/path/dirty), and coordinates "which window"
// decisions for New / Open / file-associations / second-instance launches.
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
}

const windows = new Map<number, WinEntry>()
/** Creation order, used for Ctrl+Tab / Ctrl+F6 cycling. */
const windowOrder: number[] = []
let bookCounter = 0

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
    webPreferences: { preload: join(__dirname, '../preload/index.js'), sandbox: false }
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
      return { id, title: `${entry.state.fileName} - Excel`, current: focused?.id === id }
    })
}

/** Creates a new workbook window. Exactly one of `book`/`filePath` should be given. */
function createWindow(opts: { book?: string; filePath?: string } = {}): BrowserWindow {
  const bookName = opts.book ?? (opts.filePath ? basename(opts.filePath) : nextBookName())
  let win: BrowserWindow
  let usedSpare = false

  if (!opts.filePath && spare) {
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
      webPreferences: { preload: join(__dirname, '../preload/index.js'), sandbox: false }
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
    allowClose: false
  }
  windows.set(win.id, entry)
  windowOrder.push(win.id)

  win.on('maximize', () => win.webContents.send('window:maximized', true))
  win.on('unmaximize', () => win.webContents.send('window:maximized', false))
  win.on('focus', () => broadcastWindowList())

  // Give the renderer a chance to prompt "Save changes?" before the window actually closes.
  win.on('close', (e) => {
    if (entry.allowClose) return
    e.preventDefault()
    win.webContents.send('window:closeRequested')
  })
  win.on('closed', () => {
    windows.delete(win.id)
    const idx = windowOrder.indexOf(win.id)
    if (idx >= 0) windowOrder.splice(idx, 1)
    broadcastWindowList()
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

  broadcastWindowList()
  return win
}

/** Reads `filePath` and pushes it into an already-created window's renderer. */
async function openPathInWindow(entry: WinEntry, filePath: string): Promise<void> {
  try {
    const data = await fs.readFile(filePath)
    entry.state = { fileName: basename(filePath), filePath, dirty: false }
    entry.win.webContents.send('file:openPath', { path: filePath, data: new Uint8Array(data) })
    broadcastWindowList()
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
  entry.win.setTitle(`${state.fileName} - Excel`)
  broadcastWindowList()
})

// ---------------------------------------------------------------------------
// IPC: multi-window management (View > Window group, Ctrl+Tab, Ctrl+N)
// ---------------------------------------------------------------------------

ipcMain.handle('window:new', () => {
  createWindow()
})

ipcMain.handle('window:list', () => getWindowList())

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

ipcMain.handle('file:read', async (_e, path: string) => new Uint8Array(await fs.readFile(path)))

ipcMain.handle('file:write', async (_e, path: string, data: Uint8Array) => {
  await fs.writeFile(path, data)
  return true
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

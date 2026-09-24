import { app, BrowserWindow, dialog, ipcMain } from 'electron'
import { join } from 'path'
import { promises as fs } from 'fs'

function createWindow(): void {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 800,
    minHeight: 500,
    frame: false, // Excel draws its own title bar
    backgroundColor: '#ffffff',
    title: 'Book1 - Excel',
    webPreferences: { preload: join(__dirname, '../preload/index.js'), sandbox: false }
  })
  win.on('maximize', () => win.webContents.send('window:maximized', true))
  win.on('unmaximize', () => win.webContents.send('window:maximized', false))
  if (process.env.ELECTRON_RENDERER_URL) win.loadURL(process.env.ELECTRON_RENDERER_URL)
  else win.loadFile(join(__dirname, '../renderer/index.html'))
}

ipcMain.handle('window:control', (e, action: 'minimize' | 'maximize' | 'close') => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (!win) return
  if (action === 'minimize') win.minimize()
  else if (action === 'maximize') (win.isMaximized() ? win.unmaximize() : win.maximize())
  else win.close()
})
ipcMain.handle('window:setTitle', (e, title: string) => BrowserWindow.fromWebContents(e.sender)?.setTitle(title))

const filters = [
  { name: 'Excel Workbook', extensions: ['xlsx'] },
  { name: 'CSV (Comma delimited)', extensions: ['csv'] },
  { name: 'All Files', extensions: ['*'] }
]

ipcMain.handle('file:open', async (e) => {
  const win = BrowserWindow.fromWebContents(e.sender)!
  const r = await dialog.showOpenDialog(win, { properties: ['openFile'], filters })
  if (r.canceled || !r.filePaths[0]) return null
  const path = r.filePaths[0]
  return { path, data: new Uint8Array(await fs.readFile(path)) }
})

ipcMain.handle('file:saveAs', async (e, defaultName: string) => {
  const win = BrowserWindow.fromWebContents(e.sender)!
  const r = await dialog.showSaveDialog(win, { defaultPath: defaultName, filters })
  return r.canceled ? null : r.filePath
})

ipcMain.handle('file:write', async (_e, path: string, data: Uint8Array) => {
  await fs.writeFile(path, data)
  return true
})

app.whenReady().then(createWindow)
app.on('window-all-closed', () => app.quit())

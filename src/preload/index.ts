import { contextBridge, ipcRenderer, webUtils } from 'electron'

export interface WinListEntry { id: number; title: string; current: boolean }
export interface WinState { fileName: string; filePath: string | null; dirty: boolean }
export interface OpenPathPayload { path: string; data: Uint8Array }

const api = {
  windowControl: (action: 'minimize' | 'maximize' | 'close') => ipcRenderer.invoke('window:control', action),
  setTitle: (title: string) => ipcRenderer.invoke('window:setTitle', title),
  onMaximized: (cb: (max: boolean) => void) => ipcRenderer.on('window:maximized', (_e, v) => cb(v)),

  // ---- multi-window (one workbook per window, like Excel) ----
  newWindow: (): Promise<void> => ipcRenderer.invoke('window:new'),
  listWindows: (): Promise<WinListEntry[]> => ipcRenderer.invoke('window:list'),
  activateWindow: (id: number): Promise<void> => ipcRenderer.invoke('window:activate', id),
  cycleWindow: (dir: 1 | -1): Promise<void> => ipcRenderer.invoke('window:cycle', dir),
  arrangeWindowsPrompt: (): Promise<void> => ipcRenderer.invoke('window:arrangePrompt'),
  onWindowListChanged: (cb: (list: WinListEntry[]) => void): (() => void) => {
    const listener = (_e: Electron.IpcRendererEvent, list: WinListEntry[]): void => cb(list)
    ipcRenderer.on('window:listChanged', listener)
    return () => ipcRenderer.removeListener('window:listChanged', listener)
  },
  /** Lets main know this window's current file name/path/dirty flag, for reuse & title-bar decisions. */
  reportState: (state: WinState) => ipcRenderer.send('window:reportState', state),
  /** Sent to a freshly-shown spare window to give it its "BookN" name. */
  onAssignBook: (cb: (name: string) => void) => ipcRenderer.on('window:assignBook', (_e, name) => cb(name)),

  // ---- files ----
  pickAndOpen: (): Promise<boolean> => ipcRenderer.invoke('file:pickAndOpen'),
  openPathRequest: (path: string): Promise<boolean> => ipcRenderer.invoke('file:openPathRequest', path),
  onOpenPath: (cb: (payload: OpenPathPayload) => void) => ipcRenderer.on('file:openPath', (_e, payload) => cb(payload)),
  saveAs: (defaultName: string): Promise<string | null> => ipcRenderer.invoke('file:saveAs', defaultName),
  writeFile: (path: string, data: Uint8Array): Promise<boolean> => ipcRenderer.invoke('file:write', path, data),
  readFile: (path: string): Promise<Uint8Array> => ipcRenderer.invoke('file:read', path),
  /** Resolves a dropped File to an absolute path (Electron's replacement for File.path). */
  getPathForFile: (file: File): string => webUtils.getPathForFile(file),

  /** Main asks the renderer to confirm before actually closing the window (unsaved-changes prompt). */
  onCloseRequested: (cb: () => void) => ipcRenderer.on('window:closeRequested', () => cb()),
  confirmClose: (shouldClose: boolean): Promise<void> => ipcRenderer.invoke('window:confirmClose', shouldClose),

  // ---- printing (File > Print / Ctrl+P, Page Setup, Export > Create PDF) ----
  /** Real installed printers (webContents.getPrintersAsync), for the Print pane's Printer dropdown. */
  getPrinters: (): Promise<PrinterInfo[]> => ipcRenderer.invoke('print:getPrinters'),
  /** Renders `html` (the assembled print document - print/printDoc.ts) in a hidden window and sends
   *  it to a real printer via webContents.print(). */
  printHtml: (html: string, opts: PrintHtmlOptions): Promise<boolean> => ipcRenderer.invoke('print:html', html, opts),
  /** Same, but via webContents.printToPDF (Export > Create PDF / "Microsoft Print to PDF"). */
  printToPdf: (html: string, opts: PrintToPdfOptions): Promise<Uint8Array> => ipcRenderer.invoke('print:toPdf', html, opts)
}

export interface PrinterInfo { name: string; displayName: string; description: string; isDefault: boolean }
export interface PrintHtmlOptions {
  copies: number
  collate: boolean
  deviceName?: string
  landscape: boolean
  widthMicrons: number
  heightMicrons: number
}
export interface PrintToPdfOptions { landscape: boolean; widthMicrons: number; heightMicrons: number }

contextBridge.exposeInMainWorld('api', api)
export type Api = typeof api

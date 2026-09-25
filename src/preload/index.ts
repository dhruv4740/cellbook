import { contextBridge, ipcRenderer, webUtils } from 'electron'

export interface WinListEntry { id: number; title: string; current: boolean }
export interface WinState { fileName: string; filePath: string | null; dirty: boolean }
export interface OpenPathPayload { path: string; data: Uint8Array }

const api = {
  windowControl: (action: 'minimize' | 'maximize' | 'close') => ipcRenderer.invoke('window:control', action),
  setTitle: (title: string) => ipcRenderer.invoke('window:setTitle', title),
  onMaximized: (cb: (max: boolean) => void) => ipcRenderer.on('window:maximized', (_e, v) => cb(v)),

  // ---- multi-window (one workbook per window, or multiple views of same workbook) ----
  newWindow: (opts?: { isNewWorkbook?: boolean }): Promise<void> => ipcRenderer.invoke('window:new', opts),
  newViewWindow: (snapshot: { data: unknown; fileName: string; filePath: string | null }): Promise<void> =>
    ipcRenderer.invoke('window:newView', snapshot),
  listWindows: (): Promise<WinListEntry[]> => ipcRenderer.invoke('window:list'),
  activateWindow: (id: number): Promise<void> => ipcRenderer.invoke('window:activate', id),
  cycleWindow: (dir: 1 | -1): Promise<void> => ipcRenderer.invoke('window:cycle', dir),
  arrangeWindowsPrompt: (): Promise<void> => ipcRenderer.invoke('window:arrangePrompt'),
  arrangeWindows: (mode: 'tiled' | 'horizontal' | 'vertical' | 'cascade'): Promise<void> =>
    ipcRenderer.invoke('window:arrange', mode),
  onWindowListChanged: (cb: (list: WinListEntry[]) => void): (() => void) => {
    const listener = (_e: Electron.IpcRendererEvent, list: WinListEntry[]): void => cb(list)
    ipcRenderer.on('window:listChanged', listener)
    return () => ipcRenderer.removeListener('window:listChanged', listener)
  },
  /** Lets main know this window's current file name/path/dirty flag, for reuse & title-bar decisions. */
  reportState: (state: WinState) => ipcRenderer.send('window:reportState', state),
  /** Sent to a freshly-shown spare window to give it its "BookN" name. */
  onAssignBook: (cb: (name: string) => void) => ipcRenderer.on('window:assignBook', (_e, name) => cb(name)),

  // ---- multi-window view synchronization ----
  broadcastOp: (msg: unknown): void => ipcRenderer.send('workbook:op', msg),
  onRemoteOp: (cb: (msg: unknown) => void): (() => void) => {
    const listener = (_e: Electron.IpcRendererEvent, msg: unknown): void => cb(msg)
    ipcRenderer.on('workbook:op', listener)
    return () => ipcRenderer.removeListener('workbook:op', listener)
  },
  reportWorkbookSaved: (info: { fileName: string; filePath: string | null }): void =>
    ipcRenderer.send('workbook:saved', info),
  onWorkbookSaved: (cb: (info: { fileName: string; filePath: string | null }) => void): (() => void) => {
    const listener = (_e: Electron.IpcRendererEvent, info: { fileName: string; filePath: string | null }): void => cb(info)
    ipcRenderer.on('workbook:saved', listener)
    return () => ipcRenderer.removeListener('workbook:saved', listener)
  },
  onViewIndexChanged: (cb: (viewIndex: number) => void): (() => void) => {
    const listener = (_e: Electron.IpcRendererEvent, idx: number): void => cb(idx)
    ipcRenderer.on('window:viewIndexChanged', listener)
    return () => ipcRenderer.removeListener('window:viewIndexChanged', listener)
  },
  onInitView: (cb: (payload: { data: unknown; fileName: string; filePath: string | null; viewIndex: number }) => void): (() => void) => {
    const listener = (_e: Electron.IpcRendererEvent, payload: { data: unknown; fileName: string; filePath: string | null; viewIndex: number }): void => cb(payload)
    ipcRenderer.on('window:initView', listener)
    return () => ipcRenderer.removeListener('window:initView', listener)
  },

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
  confirmClose: (shouldClose: boolean): Promise<void> => ipcRenderer.invoke('window:confirmClose', shouldClose)
}

contextBridge.exposeInMainWorld('api', api)
export type Api = typeof api

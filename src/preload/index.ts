import { contextBridge, ipcRenderer } from 'electron'

const api = {
  windowControl: (action: 'minimize' | 'maximize' | 'close') => ipcRenderer.invoke('window:control', action),
  setTitle: (title: string) => ipcRenderer.invoke('window:setTitle', title),
  onMaximized: (cb: (max: boolean) => void) => ipcRenderer.on('window:maximized', (_e, v) => cb(v)),
  openFile: (): Promise<{ path: string; data: Uint8Array } | null> => ipcRenderer.invoke('file:open'),
  saveAs: (defaultName: string): Promise<string | null> => ipcRenderer.invoke('file:saveAs', defaultName),
  writeFile: (path: string, data: Uint8Array): Promise<boolean> => ipcRenderer.invoke('file:write', path, data),
  readFile: (path: string): Promise<Uint8Array> => ipcRenderer.invoke('file:read', path),
  /** Main asks the renderer to confirm before actually closing the window (unsaved-changes prompt). */
  onCloseRequested: (cb: () => void) => ipcRenderer.on('window:closeRequested', () => cb()),
  confirmClose: (shouldClose: boolean): Promise<void> => ipcRenderer.invoke('window:confirmClose', shouldClose)
}

contextBridge.exposeInMainWorld('api', api)
export type Api = typeof api

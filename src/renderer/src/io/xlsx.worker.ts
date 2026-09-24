// Runs the fast xlsx parser off the main thread so opening a big workbook never freezes the UI.
// Posts { ok: true, data } on success or { ok: false, error } on failure; xlsx.ts (readXlsx)
// owns spawning this worker and falls back to parsing inline when Worker isn't available.
import { parseXlsx } from './fastXlsx'
import type { WorkbookData } from '../model/store'

self.onmessage = (ev: MessageEvent<ArrayBuffer>): void => {
  try {
    const bytes = new Uint8Array(ev.data)
    const data: WorkbookData = parseXlsx(bytes)
    self.postMessage({ ok: true, data })
  } catch (err) {
    self.postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) })
  }
}

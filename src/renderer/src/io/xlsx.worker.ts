// Runs the fast xlsx parser off the main thread so opening a big workbook never freezes the UI.
// Posts { ok: true, data } on success or { ok: false, error } on failure; xlsx.ts (readXlsx)
// owns spawning this worker and falls back to parsing inline when Worker isn't available.
import { parseXlsx } from './fastXlsx'
import { parseXls } from './xls/biff8'
import { isCfbMagic } from './xlsxShared'

self.onmessage = (ev: MessageEvent<ArrayBuffer>): void => {
  try {
    const bytes = new Uint8Array(ev.data)
    // vbaProject.bin bytes are structured-cloned (not transferred) by postMessage, same as the
    // rest of `data` - fine for the one-off cost of an open, and simplest to reason about.
    // Detected by CFB magic (D0 CF 11 E0), not by file extension: a renamed .xls still opens.
    const data = isCfbMagic(bytes) ? parseXls(bytes) : parseXlsx(bytes)
    self.postMessage({ ok: true, data })
  } catch (err) {
    self.postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) })
  }
}

// Runs the fast xlsx parser off the main thread so opening a big workbook never freezes the UI.
// Posts { ok: true, data } on success or { ok: false, error } on failure; xlsx.ts (readXlsx)
// owns spawning this worker and falls back to parsing inline when Worker isn't available.
import { parseXlsx } from './fastXlsx'
import { packCells } from './packedCells'

/** Sheets with at least this many cells travel in packed form (typed arrays are transferred, not cloned). */
const PACK_MIN = 20000

self.onmessage = (ev: MessageEvent<ArrayBuffer>): void => {
  try {
    const bytes = new Uint8Array(ev.data)
    const data = parseXlsx(bytes)
    const transfer: ArrayBuffer[] = []
    for (const sheet of data.sheets) {
      if (Object.keys(sheet.cells).length < PACK_MIN) continue
      const packed = packCells(sheet.cells)
      sheet.packed = packed
      sheet.cells = {}
      transfer.push(packed.r.buffer as ArrayBuffer, packed.c.buffer as ArrayBuffer, packed.sid.buffer as ArrayBuffer, packed.len.buffer as ArrayBuffer)
    }
    // vbaProject.bin bytes are structured-cloned (not transferred), same as the rest of `data`.
    ;(self as unknown as Worker).postMessage({ ok: true, data }, transfer)
  } catch (err) {
    self.postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) })
  }
}

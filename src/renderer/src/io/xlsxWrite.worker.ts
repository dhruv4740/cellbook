// Builds the .xlsx bytes off the main thread so saving a big workbook never freezes the UI.
import { writeXlsx } from './xlsx'
import type { WorkbookData } from '../model/store'
import type { VbaProject } from './xlsxPostProcess'

self.onmessage = async (ev: MessageEvent<{ data: WorkbookData; vba?: VbaProject }>): Promise<void> => {
  try {
    const bytes = await writeXlsx(ev.data.data, ev.data.vba)
    ;(self as unknown as Worker).postMessage({ ok: true, bytes }, [bytes.buffer as ArrayBuffer])
  } catch (err) {
    self.postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) })
  }
}

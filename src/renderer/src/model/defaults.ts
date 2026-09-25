// Dependency-free default values shared by the heavy workbook store (HyperFormula/React) and the
// fast, worker-side xlsx reader (io/fastXlsx.ts, io/xlsx.worker.ts), which must not pull those in.
// store.ts re-exports these so existing `from '../model/store'` imports keep working.
import type { SheetLayout } from './types'
import { defaultPageSetup } from './types'

/** Default author shown on new notes/comments (Excel shows the signed-in account's name; this
 *  app has none, so it falls back to the OS-style default Excel itself uses when unlicensed). */
export const DEFAULT_AUTHOR = 'Excel User'

export function defaultLayout(): SheetLayout {
  return {
    colWidths: {}, rowHeights: {}, hiddenRows: {}, hiddenCols: {}, merges: [],
    frozen: { rows: 0, cols: 0 }, showGridlines: true, showHeadings: true, zoom: 100,
    notes: {}, comments: {}, pageSetup: defaultPageSetup(), tables: []
  }
}

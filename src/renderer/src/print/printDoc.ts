// Assembles the full print document: every page of every sheet being printed, as one HTML string
// with @media print page-break rules. Shared by the Print backstage preview (srcdoc of an iframe),
// the hidden Electron print BrowserWindow, printToPDF export and the browser window.print() fallback.
import { PageSetup } from '../model/types'
import { WorkbookStore, SheetState } from '../model/store'
import { computePrintPlan, PrintPage } from './paginate'
import { pageGeometry } from './pageGeometry'
import { renderPageHtml } from './renderPage'

export interface PrintTarget {
  sheet: SheetState
  pageSetup: PageSetup
}

export interface PrintDocResult {
  html: string
  pageCount: number
  /** Paper size in px of the first sheet printed (pages can vary per-sheet when printing the whole workbook). */
  widthPx: number
  heightPx: number
}

const PRINT_CSS = `
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #808080; }
.xl-print-page {
  background: #fff;
  margin: 0 auto 12px auto;
  position: relative;
  overflow: hidden;
  box-shadow: 0 0 6px rgba(0,0,0,.35);
  display: flex;
  flex-direction: column;
}
.xl-print-header, .xl-print-footer { display: flex; align-items: center; font: 9pt Calibri, sans-serif; color: #000; flex-shrink: 0; }
.xl-print-footer { align-items: flex-end; }
.xl-hf-l { flex: 1; text-align: left; }
.xl-hf-c { flex: 1; text-align: center; }
.xl-hf-r { flex: 1; text-align: right; }
.xl-print-body { flex: 1; overflow: hidden; }
.xl-print-table { border-collapse: collapse; }
.xl-print-table td { padding: 0 2px; font: 11pt Calibri, sans-serif; overflow: hidden; }
.xl-print-table.xl-print-gridlines td { border: 1px solid #d0d0d0; }
.xl-print-colhead, .xl-print-rowhead, .xl-print-corner {
  background: #f0f0f0; border: 1px solid #999; text-align: center; font: 8pt Calibri, sans-serif; color: #333;
}
@media print {
  html, body { background: #fff; }
  .xl-print-page { box-shadow: none; margin: 0; page-break-after: always; }
  .xl-print-page:last-child { page-break-after: auto; }
}
`

/** Builds the print document for one or more sheets (File > Print's "Entire Workbook" prints every
 *  sheet's own pages back to back; "Print Active Sheets"/"Selection" pass a single target). */
export function buildPrintDoc(store: WorkbookStore, targets: PrintTarget[], fileName: string): PrintDocResult {
  const now = new Date()
  let html = ''
  let pageCount = 0
  let widthPx = 0, heightPx = 0
  const perSheetPages: { sheet: SheetState; ps: PageSetup; pages: PrintPage[]; scalePercent: number }[] = []

  for (const t of targets) {
    const plan = computePrintPlan(store, t.sheet, t.pageSetup)
    perSheetPages.push({ sheet: t.sheet, ps: t.pageSetup, pages: plan.pages, scalePercent: plan.scalePercent })
    pageCount += plan.pages.length
    const geo = pageGeometry(t.pageSetup)
    widthPx = widthPx || geo.paperWidthPx
    heightPx = heightPx || geo.paperHeightPx
  }

  let globalPage = 0
  for (const { sheet, ps, pages, scalePercent } of perSheetPages) {
    for (const page of pages) {
      globalPage++
      html += renderPageHtml(store, sheet, ps, {
        page,
        pageCount,
        scalePercent,
        hfCtx: { date: now, fileName, sheetName: sheet.name, filePath: fileName }
      })
    }
  }

  if (!pageCount) {
    html = `<div class="xl-print-page" style="width:816px;height:1056px;display:flex;align-items:center;justify-content:center;color:#888;font:14pt Calibri">Nothing to print</div>`
  }

  const doc = `<!doctype html><html><head><meta charset="utf-8"><style>${PRINT_CSS}</style></head><body>${html}</body></html>`
  return { html: doc, pageCount: pageCount || 1, widthPx: widthPx || 816, heightPx: heightPx || 1056 }
}

/** One page's HTML for the Print backstage preview (1-based `pageIndex`), so the preview iframe
 *  only ever renders the page currently shown instead of the whole (possibly huge) document. */
export function buildSinglePageDoc(store: WorkbookStore, sheet: SheetState, ps: PageSetup, pageIndex: number, fileName: string): PrintDocResult {
  const plan = computePrintPlan(store, sheet, ps)
  const geo = pageGeometry(ps)
  if (!plan.pages.length) {
    return {
      html: `<!doctype html><html><head><meta charset="utf-8"><style>${PRINT_CSS}</style></head><body><div class="xl-print-page" style="width:${geo.paperWidthPx}px;height:${geo.paperHeightPx}px;display:flex;align-items:center;justify-content:center;color:#888;font:14pt Calibri">Nothing to print</div></body></html>`,
      pageCount: 1,
      widthPx: geo.paperWidthPx,
      heightPx: geo.paperHeightPx
    }
  }
  const page = plan.pages[Math.max(0, Math.min(pageIndex, plan.pages.length - 1))]
  const html = renderPageHtml(store, sheet, ps, {
    page,
    pageCount: plan.pages.length,
    scalePercent: plan.scalePercent,
    hfCtx: { date: new Date(), fileName, sheetName: sheet.name, filePath: fileName }
  })
  const doc = `<!doctype html><html><head><meta charset="utf-8"><style>${PRINT_CSS}</style></head><body>${html}</body></html>`
  return { html: doc, pageCount: plan.pages.length, widthPx: geo.paperWidthPx, heightPx: geo.paperHeightPx }
}

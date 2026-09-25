// Actually printing: in Electron this hands the assembled HTML to the main process, which loads it
// into a hidden BrowserWindow and calls webContents.print()/printToPDF(); in a plain browser
// (npm run web) it falls back to window.print() on a hidden iframe holding the same HTML, so
// preview and the real output never diverge.
import { WorkbookStore, SheetState } from '../model/store'
import { PageSetup } from '../model/types'
import { buildPrintDoc, PrintTarget } from './printDoc'

export interface PrintJobOptions {
  copies: number
  collate: boolean
  deviceName?: string // '' / undefined = system default
  landscape: boolean
  paperWidthPx: number
  paperHeightPx: number
}

function browserFallbackPrint(html: string): void {
  const iframe = document.createElement('iframe')
  iframe.style.position = 'fixed'
  iframe.style.right = '0'
  iframe.style.bottom = '0'
  iframe.style.width = '0'
  iframe.style.height = '0'
  iframe.style.border = '0'
  document.body.appendChild(iframe)
  iframe.onload = () => {
    iframe.contentWindow?.focus()
    iframe.contentWindow?.print()
    setTimeout(() => iframe.remove(), 1000)
  }
  iframe.srcdoc = html
}

export async function printTargets(store: WorkbookStore, targets: PrintTarget[], fileName: string, opts: PrintJobOptions): Promise<void> {
  const doc = buildPrintDoc(store, targets, fileName)
  if (window.api?.printHtml) {
    await window.api.printHtml(doc.html, {
      copies: opts.copies,
      collate: opts.collate,
      deviceName: opts.deviceName,
      landscape: opts.landscape,
      widthMicrons: Math.round((opts.paperWidthPx / 96) * 25400),
      heightMicrons: Math.round((opts.paperHeightPx / 96) * 25400)
    })
  } else {
    browserFallbackPrint(doc.html)
  }
}

/** File > Export > Create PDF (Electron's printToPDF; falls back to window.print() -> "Save as
 *  PDF" in a browser, since a plain web page can't drive the OS print-to-PDF pipeline itself). */
export async function exportTargetsToPdf(store: WorkbookStore, targets: PrintTarget[], fileName: string): Promise<Uint8Array | null> {
  const doc = buildPrintDoc(store, targets, fileName)
  if (window.api?.printToPdf) {
    return window.api.printToPdf(doc.html, {
      landscape: targets[0]?.pageSetup.orientation === 'landscape',
      widthMicrons: Math.round((doc.widthPx / 96) * 25400),
      heightMicrons: Math.round((doc.heightPx / 96) * 25400)
    })
  }
  browserFallbackPrint(doc.html)
  return null
}

export function singleTargetsForActiveSheet(sheet: SheetState, override?: Partial<PageSetup>): PrintTarget[] {
  const ps = { ...(sheet.layout.pageSetup!), ...override }
  return [{ sheet, pageSetup: ps }]
}

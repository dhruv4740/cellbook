// Page Layout > Background: a tiled image painted behind a sheet's cells (screen only, never printed).
import { store, SheetState } from '../model/store'

/** Sets (or, with null, deletes) the sheet's background as one undo step. */
export function setSheetBackground(sheet: SheetState, src: string | null): void {
  if (store.refuseIfProtected(sheet)) return
  store.updateLayout(sheet, src ? 'Sheet Background' : 'Delete Background', (l) => {
    if (src) l.background = src
    else delete l.background
  })
}

/** Opens the image picker and uses the chosen file as the active sheet's background. */
export function pickSheetBackground(): void {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = 'image/png,image/jpeg,image/gif,image/bmp,image/webp'
  input.onchange = async () => {
    const f = input.files?.[0]
    if (!f) return
    const src = await new Promise<string>((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result)); fr.onerror = () => rej(fr.error); fr.readAsDataURL(f) })
    setSheetBackground(store.active, src)
  }
  input.click()
}

/** CSS custom properties the grid host uses to tile the background under the cell area. */
export function backgroundVars(src: string | undefined, rowHeaderW: number, headerH: number, scroll: { x: number; y: number }): Record<string, string> {
  if (!src) return {}
  return {
    '--xl-bg-img': `url("${src}")`,
    '--xl-bg-x': `${rowHeaderW}px`,
    '--xl-bg-y': `${headerH}px`,
    '--xl-bg-sx': `${-scroll.x}px`,
    '--xl-bg-sy': `${-scroll.y}px`
  }
}

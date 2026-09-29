// Insert > Pictures: floating images on a sheet. Every change goes through store.updateLayout so it undoes.
import { store, SheetState } from '../model/store'
import { PictureDef } from '../model/types'
import { setUi } from '../chrome/ui'

const MAX_W = 800, MAX_H = 600

export const picturesOf = (sheet: SheetState): PictureDef[] => sheet.layout.pictures ?? []

function measure(src: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight })
    img.onerror = () => reject(new Error('unreadable image'))
    img.src = src
  })
}

/** Puts an image at the active cell (native size, shrunk to fit a 800 x 600 box like Excel keeps it on screen). */
export async function insertPictureFromDataUrl(src: string, _fileName?: string): Promise<void> {
  const sheet = store.active
  if (!store.allows(sheet, 'objects')) return
  let size: { w: number; h: number }
  try { size = await measure(src) } catch { setUi({ dialog: { kind: 'message', props: { text: 'We can’t insert that picture. It may be damaged or in a format that isn’t supported.' } } }); return }
  const k = Math.min(1, MAX_W / size.w, MAX_H / size.h)
  const { r, c } = sheet.selection.active
  const n = picturesOf(sheet).length + 1
  const pic: PictureDef = {
    id: 'pic_' + Date.now().toString(36) + n, name: `Picture ${n}`,
    src, anchor: { r, c, dx: 0, dy: 0 }, width: Math.max(8, Math.round(size.w * k)), height: Math.max(8, Math.round(size.h * k))
  }
  store.updateLayout(sheet, 'Insert Picture', (l) => { l.pictures = [...(l.pictures ?? []), pic] })
  setUi({ selectedPicture: pic.id, selectedChart: null })
}

/** Opens the file picker and inserts the chosen image(s). */
export function pickPictureFiles(): void {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = 'image/png,image/jpeg,image/gif,image/bmp,image/webp'
  input.multiple = true
  input.onchange = async () => {
    for (const f of Array.from(input.files ?? [])) {
      const src = await new Promise<string>((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result)); fr.onerror = () => rej(fr.error); fr.readAsDataURL(f) })
      await insertPictureFromDataUrl(src, f.name)
    }
  }
  input.click()
}

export function updatePicture(sheet: SheetState, id: string, patch: Partial<PictureDef>, label = 'Format Picture'): void {
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, label, (l) => { l.pictures = (l.pictures ?? []).map((p) => (p.id === id ? { ...p, ...patch } : p)) })
}

export function deletePicture(sheet: SheetState, id: string): void {
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, 'Delete Picture', (l) => { l.pictures = (l.pictures ?? []).filter((p) => p.id !== id) })
}

export function reorderPicture(sheet: SheetState, id: string, toFront: boolean): void {
  if (!store.allows(sheet, 'objects')) return
  store.updateLayout(sheet, toFront ? 'Bring to Front' : 'Send to Back', (l) => {
    const list = l.pictures ?? []
    const p = list.find((x) => x.id === id)
    if (p) l.pictures = toFront ? [...list.filter((x) => x.id !== id), p] : [p, ...list.filter((x) => x.id !== id)]
  })
}

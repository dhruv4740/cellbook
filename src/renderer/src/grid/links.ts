// Cell hyperlinks: following them (click) and inserting / removing them (Ctrl+K).
import { store, SheetState } from '../model/store'
import { parseA1, HyperlinkData } from '../model/types'
import { setUi, showToast } from '../chrome/ui'

export function linkAt(sheet: SheetState, r: number, c: number): HyperlinkData | undefined {
  return sheet.layout.hyperlinks?.[`${r},${c}`]
}

/** Opens a link like Excel: web/mail/file targets outside, "Sheet!A1" locations inside the workbook. */
export async function followLink(link: HyperlinkData): Promise<void> {
  if (link.location) {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(link.location.replace(/^#/, ''))
    const sheetName = m ? (m[1] ?? m[2]).replace(/''/g, "'") : store.active.name
    const ref = (m ? m[3] : link.location.replace(/^#/, '')).replace(/\$/g, '')
    const idx = store.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase())
    const rg = parseA1(ref)
    if (idx < 0 || !rg) { setUi({ dialog: { kind: 'message', props: { text: 'Reference isn\u2019t valid.' } } }); return }
    store.setActiveSheet(idx)
    store.setSelection({ ranges: [rg], active: { r: rg.r1, c: rg.c1 } }, store.sheets[idx])
    return
  }
  if (!link.target) return
  if (!window.api?.openLink) { window.open(link.target, '_blank', 'noopener'); return }
  const res = await window.api.openLink(link.target)
  if (res !== true) showToast(`Cannot open the specified file. ${res}`)
}

/** Insert Link dialog OK: sets the link (and the display text), one undo step. */
export function setLink(sheet: SheetState, r: number, c: number, link: HyperlinkData | null, text?: string): void {
  if (!store.allows(sheet, 'insertHyperlinks')) return
  store.transact(link ? 'Insert Hyperlink' : 'Remove Hyperlink', () => {
    if (text !== undefined && text !== store.getRaw(sheet, r, c) && !store.getRaw(sheet, r, c).startsWith('=')) store.setRaw(sheet, r, c, text)
    store.updateLayout(sheet, link ? 'Insert Hyperlink' : 'Remove Hyperlink', (l) => {
      const all = { ...(l.hyperlinks ?? {}) }
      if (link) all[`${r},${c}`] = link
      else delete all[`${r},${c}`]
      l.hyperlinks = all
    })
    const st = store.getStyle(sheet, r, c)
    if (link && !st?.color) store.applyStyle(sheet, [{ r1: r, c1: c, r2: r, c2: c }], { color: '#0563C1', underline: 'single' }, 'Hyperlink')
    if (!link && st?.color === '#0563C1') store.applyStyle(sheet, [{ r1: r, c1: c, r2: r, c2: c }], { color: undefined, underline: undefined }, 'Hyperlink')
  })
}

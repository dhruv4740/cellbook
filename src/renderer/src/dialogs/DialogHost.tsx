// Renders whichever dialog ui.dialog names, and owns the dialog keyboard shortcuts.
// Kept separate from App so store changes re-render only this component, not the whole window.
import { useEffect } from 'react'
import { setUi, useUi } from '../chrome/ui'
import { useStore } from '../model/store'
import { FormatCellsDialog } from './FormatCellsDialog'
import { FindReplaceDialog } from './FindReplaceDialog'
import { GoToDialog } from './GoToDialog'
import { GoToSpecialDialog } from './GoToSpecialDialog'

export default function DialogHost() {
  const ui = useUi()
  const store = useStore()

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey
      const k = e.key.toLowerCase()
      if (mod && !e.shiftKey && k === 'f') setUi({ dialog: { kind: 'find' } })
      else if (mod && !e.shiftKey && k === 'h') setUi({ dialog: { kind: 'replace' } })
      else if ((e.key === 'F5' && !mod) || (mod && k === 'g')) setUi({ dialog: { kind: 'goto' } })
      else return
      e.preventDefault()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  const d = ui.dialog
  if (!d) return null
  const close = () => setUi({ dialog: null })
  const sheet = store.active
  const { r, c } = sheet.selection.active

  switch (d.kind) {
    case 'formatCells':
      return (
        <FormatCellsDialog
          initialTab={(d.props?.tab as string) || 'Number'}
          initialStyle={store.getStyle(sheet, r, c)}
          cellValue={store.getValue(sheet, r, c)}
          onApply={(patch) => store.applyStyle(sheet, sheet.selection.ranges, patch)}
          onClose={close}
        />
      )
    case 'find':
    case 'replace':
      return <FindReplaceDialog initialTab={d.kind} store={store} onClose={close} />
    case 'goto':
      return <GoToDialog onOpenSpecial={() => setUi({ dialog: { kind: 'gotoSpecial' } })} onClose={close} />
    case 'gotoSpecial':
      return <GoToSpecialDialog store={store} initialType={(d.props?.type as never) || 'formulas'} onClose={close} />
    default:
      return null
  }
}

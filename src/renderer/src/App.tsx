// Top-level window layout, top to bottom, exactly like Excel.
import { useEffect } from 'react'
import TitleBar from './chrome/TitleBar'
import Ribbon from './chrome/Ribbon'
import FormulaBar from './chrome/FormulaBar'
import Grid from './grid/Grid'
import SheetTabs from './chrome/SheetTabs'
import StatusBar from './chrome/StatusBar'
import Backstage from './io/Backstage'
import { setUi, useUi } from './chrome/ui'
import { useStore } from './model/store'
import { FindReplaceDialog } from './dialogs/FindReplaceDialog'
import { GoToDialog } from './dialogs/GoToDialog'
import { GoToSpecialDialog } from './dialogs/GoToSpecialDialog'

export default function App() {
  const ui = useUi()
  const store = useStore()

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey
      if (mod && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault()
        setUi({ dialog: { kind: 'find' } })
      } else if (mod && (e.key === 'h' || e.key === 'H')) {
        e.preventDefault()
        setUi({ dialog: { kind: 'replace' } })
      } else if (e.key === 'F5' || (mod && (e.key === 'g' || e.key === 'G'))) {
        e.preventDefault()
        setUi({ dialog: { kind: 'goto' } })
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [])

  if (ui.backstage) return <Backstage />
  return (
    <div className="app">
      <TitleBar />
      <Ribbon />
      <FormulaBar />
      <Grid />
      <SheetTabs />
      <StatusBar />

      {(ui.dialog?.kind === 'find' || ui.dialog?.kind === 'replace') && (
        <FindReplaceDialog
          initialTab={ui.dialog.kind}
          store={store}
          onClose={() => setUi({ dialog: null })}
        />
      )}

      {ui.dialog?.kind === 'goto' && (
        <GoToDialog
          onOpenSpecial={() => setUi({ dialog: { kind: 'gotoSpecial' } })}
          onClose={() => setUi({ dialog: null })}
        />
      )}

      {ui.dialog?.kind === 'gotoSpecial' && (
        <GoToSpecialDialog
          store={store}
          initialType={(ui.dialog.props?.type as any) || 'formulas'}
          onClose={() => setUi({ dialog: null })}
        />
      )}
    </div>
  )
}

// Top-level window layout, top to bottom, exactly like Excel.
import { useEffect } from 'react'
import TitleBar from './chrome/TitleBar'
import Ribbon from './chrome/Ribbon'
import FormulaBar from './chrome/FormulaBar'
import Grid from './grid/Grid'
import SheetTabs from './chrome/SheetTabs'
import StatusBar from './chrome/StatusBar'
import Backstage from './io/Backstage'
import DialogHost from './dialogs/DialogHost'
import { useUi } from './chrome/ui'
import { installFileShortcuts } from './io/fileOps'

export default function App() {
  const ui = useUi()
  useEffect(() => {
    installFileShortcuts()
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
      <DialogHost />
      {ui.toast && <div className="xl-toast" role="status">{ui.toast}</div>}
    </div>
  )
}

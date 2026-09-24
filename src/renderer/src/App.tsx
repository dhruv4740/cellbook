// Top-level window layout, top to bottom, exactly like Excel.
import TitleBar from './chrome/TitleBar'
import Ribbon from './chrome/Ribbon'
import FormulaBar from './chrome/FormulaBar'
import Grid from './grid/Grid'
import SheetTabs from './chrome/SheetTabs'
import StatusBar from './chrome/StatusBar'
import Backstage from './io/Backstage'
import { useUi } from './chrome/ui'

export default function App() {
  const ui = useUi()
  if (ui.backstage) return <Backstage />
  return (
    <div className="app">
      <TitleBar />
      <Ribbon />
      <FormulaBar />
      <Grid />
      <SheetTabs />
      <StatusBar />
    </div>
  )
}

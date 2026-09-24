// Top-level window layout, top to bottom, exactly like Excel.
import TitleBar from './chrome/TitleBar'
import Ribbon from './chrome/Ribbon'
import FormulaBar from './chrome/FormulaBar'
import Grid from './grid/Grid'
import SheetTabs from './chrome/SheetTabs'
import StatusBar from './chrome/StatusBar'
import Backstage from './io/Backstage'
import { setUi, useUi } from './chrome/ui'
import { useStore } from './model/store'
import { FormatCellsDialog } from './dialogs/FormatCellsDialog'

export default function App() {
  const ui = useUi()
  const store = useStore()

  if (ui.backstage) return <Backstage />

  const activeSheet = store.active
  const activeCell = activeSheet?.selection?.active || { r: 0, c: 0 }
  const activeStyle = activeSheet ? store.getStyle(activeSheet, activeCell.r, activeCell.c) : undefined
  const activeValue = activeSheet ? store.getValue(activeSheet, activeCell.r, activeCell.c) : undefined

  return (
    <div className="app">
      <TitleBar />
      <Ribbon />
      <FormulaBar />
      <Grid />
      <SheetTabs />
      <StatusBar />
      {ui.dialog?.kind === 'formatCells' && (
        <FormatCellsDialog
          initialTab={(ui.dialog.props?.tab as string) || 'Number'}
          initialStyle={activeStyle}
          cellValue={activeValue}
          onApply={(patch) => {
            if (activeSheet) {
              store.applyStyle(activeSheet, activeSheet.selection.ranges, patch)
            }
          }}
          onClose={() => setUi({ dialog: null })}
        />
      )}
    </div>
  )
}

// Renders whichever dialog ui.dialog names, and owns the dialog keyboard shortcuts.
// Kept separate from App so store changes re-render only this component, not the whole window.
import { useEffect } from 'react'
import { setUi, useUi } from '../chrome/ui'
import { useStore } from '../model/store'
import { FormatCellsDialog } from './FormatCellsDialog'
import { FindReplaceDialog } from './FindReplaceDialog'
import { GoToDialog } from './GoToDialog'
import { GoToSpecialDialog } from './GoToSpecialDialog'
import { CfQuickDialog, CfQuickKind } from './CfQuickDialog'
import { CfNewRuleDialog } from './CfNewRuleDialog'
import { CfManageDialog } from './CfManageDialog'
import { CfRule } from '../model/types'
import { SortDialog } from './SortDialog'
import { SortWarningDialog } from './SortWarningDialog'
import { RemoveDuplicatesDialog } from './RemoveDuplicatesDialog'
import { TextToColumnsDialog } from './TextToColumnsDialog'
import { DataValidationDialog } from './DataValidationDialog'
import { DataValidationErrorDialog } from './DataValidationErrorDialog'
import { PageSetupDialog, PageSetupTab } from './PageSetupDialog'
import { CreateTableDialog } from './CreateTableDialog'
import { ResizeTableDialog } from './ResizeTableDialog'
import { RenameTableDialog } from './RenameTableDialog'
import { Dialog } from '../components/Dialog'
import { Table, Range } from '../model/types'
import { convertToRange } from '../model/tables'

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
    case 'cfQuick':
      return <CfQuickDialog kind={(d.props?.kind as CfQuickKind) || 'greaterThan'} store={store} onClose={close} />
    case 'cfNewRule':
      return <CfNewRuleDialog store={store} editRule={d.props?.rule as CfRule | undefined} onClose={close} />
    case 'cfManage':
      return (
        <CfManageDialog
          store={store}
          onClose={close}
          onNewRule={() => setUi({ dialog: { kind: 'cfNewRule' } })}
          onEditRule={(rule) => setUi({ dialog: { kind: 'cfNewRule', props: { rule } } })}
        />
      )
    case 'sort':
      return <SortDialog store={store} onClose={close} />
    case 'sortWarning':
      return (
        <SortWarningDialog
          onSort={(expand) => {
            const onSort = d.props?.onSort as ((expand: boolean) => void) | undefined
            if (onSort) onSort(expand)
          }}
          onClose={close}
        />
      )
    case 'removeDuplicates':
      return <RemoveDuplicatesDialog store={store} onClose={close} />
    case 'textToColumns':
      return <TextToColumnsDialog store={store} onClose={close} />
    case 'dataValidation':
      return <DataValidationDialog store={store} onClose={close} />
    case 'pageSetup':
      return <PageSetupDialog store={store} sheet={sheet} initialTab={d.props?.tab as PageSetupTab | undefined} onClose={close} />
    case 'createTable':
      return <CreateTableDialog store={store} range={d.props?.range as Range} styleName={d.props?.styleName as string | undefined} onClose={close} />
    case 'resizeTable':
      return <ResizeTableDialog store={store} table={d.props?.table as Table} onClose={close} />
    case 'renameTable':
      return <RenameTableDialog store={store} table={d.props?.table as Table} onClose={close} />
    case 'convertTableToRange':
      return (
        <Dialog title="Microsoft Excel" width={360} onOk={() => { convertToRange(store, sheet, d.props?.table as Table); close() }} onCancel={close} okLabel="Yes" cancelLabel="No">
          <p style={{ margin: 0, font: '12px var(--ui-font, "Segoe UI", Tahoma, sans-serif)' }}>Do you want to convert the table to a normal range?</p>
        </Dialog>
      )
    case 'dataValidationError':
      return (
        <DataValidationErrorDialog
          style={d.props?.style as any}
          title={d.props?.title as string}
          message={d.props?.message as string}
          onRetry={((d.props?.onRetry as () => void) || close)}
          onCancel={((d.props?.onCancel as () => void) || close)}
          onContinue={((d.props?.onContinue as () => void) || close)}
        />
      )
    default:
      return null
  }
}

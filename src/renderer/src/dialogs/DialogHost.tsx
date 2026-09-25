// Renders whichever dialog ui.dialog names, and owns the dialog keyboard shortcuts.
// Kept separate from App so store changes re-render only this component, not the whole window.
import { useEffect, useRef } from 'react'
import { setUi, useUi, getUi } from '../chrome/ui'
import { formatAsTable } from '../grid/commands'
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
import { ArrangeWindowsDialog } from './ArrangeWindowsDialog'
import { DefineNameDialog, NameManagerDialog, CreateFromSelectionDialog } from './NameDialogs'
import { HyperlinkDialog } from './HyperlinkDialog'
import { SubtotalDialog } from './SubtotalDialog'
import { MessageDialog, UnhideSheetDialog, MoveCopySheetDialog, SeriesDialog } from './SheetDialogs'

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
      else if (mod && !e.shiftKey && !e.altKey && k === 'k') setUi({ dialog: { kind: 'hyperlink' } })
      else if (mod && !e.shiftKey && !e.altKey && (k === 't' || k === 'l') && !getUi().edit) formatAsTable(0)
      else if (mod && e.key === 'F3') setUi({ dialog: { kind: e.shiftKey ? 'createNames' : 'nameManager' } })
      else return
      e.preventDefault()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  // When a dialog closes, give the keyboard back to the sheet (Excel), unless something else took focus.
  const hadDialog = useRef(false)
  useEffect(() => {
    if (ui.dialog) { hadDialog.current = true; return }
    if (!hadDialog.current) return
    hadDialog.current = false
    requestAnimationFrame(() => {
      const a = document.activeElement
      if (!a || a === document.body || !document.body.contains(a)) (document.querySelector('.grid-host') as HTMLElement | null)?.focus()
    })
  }, [ui.dialog])

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
    case 'arrangeWindows':
      return <ArrangeWindowsDialog onClose={close} />
    case 'message':
      return <MessageDialog text={String(d.props?.text ?? '')} onClose={close} />
    case 'defineName':
      return <DefineNameDialog onClose={close} />
    case 'nameManager':
      return <NameManagerDialog onClose={close} />
    case 'createNames':
      return <CreateFromSelectionDialog onClose={close} />
    case 'subtotal':
      return <SubtotalDialog onClose={close} />
    case 'hyperlink':
      return <HyperlinkDialog onClose={close} />
    case 'series':
      return <SeriesDialog onClose={close} />
    case 'unhideSheet':
      return <UnhideSheetDialog onClose={close} />
    case 'moveCopySheet':
      return <MoveCopySheetDialog index={(d.props?.index as number | undefined) ?? store.activeIndex} onClose={close} />
    default:
      return null
  }
}

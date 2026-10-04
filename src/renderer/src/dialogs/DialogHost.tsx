// Renders whichever dialog ui.dialog names, and owns the dialog keyboard shortcuts.
// Kept separate from App so store changes re-render only this component, not the whole window.
import { PageSetupDialog } from './PageSetupDialog'
import { AdvancedFilterDialog, GoalSeekDialog, GoalSeekStatusDialog, ConsolidateDialog } from './DataToolDialogs'
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
import { DefineNameDialog, NameManagerDialog, CreateFromSelectionDialog, ApplyNamesDialog } from './NameDialogs'
import { HyperlinkDialog } from './HyperlinkDialog'
import { SparklineDialog } from './SparklineDialog'
import { QuickAnalysisDialog } from './QuickAnalysisDialog'
import { SpellingDialog } from './SpellingDialog'
import { EvaluateFormulaDialog } from './EvaluateFormulaDialog'
import { SymbolDialog } from './SymbolDialog'
import { IconsDialog } from './IconsDialog'
import { EditRangesDialog } from './EditRangesDialog'
import { NewTableStyleDialog } from './NewTableStyleDialog'
import { AddPenDialog } from './AddPenDialog'
import { CustomViewsDialog } from './CustomViewsDialog'
import { AccessibilityDialog } from './AccessibilityDialog'
import { DataTableDialog, ScenarioManagerDialog } from './WhatIfDialogs'
import { SubtotalDialog } from './SubtotalDialog'
import { PivotDialog } from './PivotDialog'
import { ForecastDialog } from './ForecastDialog'
import { OutlineAxisDialog, OutlineSettingsDialog } from './OutlineDialogs'
import { PasswordPromptDialog, ProtectSheetDialog, ProtectWorkbookDialog } from './ProtectDialogs'
import { MessageDialog, UnhideSheetDialog, MoveCopySheetDialog, SeriesDialog } from './SheetDialogs'

export default function DialogHost() {
  const ui = useUi()
  const store = useStore()

  // Excel's "protected sheet" message whenever the store refuses an edit.
  useEffect(() => {
    store.onProtectedEdit = (text) => setUi({ dialog: { kind: 'message', props: { text } } })
    return () => { store.onProtectedEdit = undefined }
  }, [store])

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
    case 'outlineAxis':
      return <OutlineAxisDialog mode={d.props?.mode as never} onClose={close} />
    case 'outlineSettings':
      return <OutlineSettingsDialog onClose={close} />
    case 'protectSheet':
      return <ProtectSheetDialog sheetUid={(d.props?.sheetUid as number | undefined) ?? sheet.uid} onClose={close} />
    case 'advancedFilter':
      return <AdvancedFilterDialog onClose={close} />
    case 'goalSeek':
      return <GoalSeekDialog onClose={close} />
    case 'goalSeekStatus':
      return <GoalSeekStatusDialog text={String(d.props?.text)} target={String(d.props?.target)} current={String(d.props?.current)} onOk={close} onCancel={() => { close(); (d.props?.onCancel as () => void)() }} />
    case 'consolidate':
      return <ConsolidateDialog onClose={close} />
    case 'pageSetup':
      return <PageSetupDialog initialTab={d.props?.tab as never} onClose={close} />
    case 'protectWorkbook':
      return <ProtectWorkbookDialog onClose={close} />
    case 'password':
      return <PasswordPromptDialog title={String(d.props?.title)} verify={d.props?.verify as (pw: string) => boolean} onOk={d.props?.onOk as () => void} onClose={close} />
    case 'accessibility':
      return <AccessibilityDialog onClose={close} />
    case 'applyNames':
      return <ApplyNamesDialog onClose={close} />
    case 'defineName':
      return <DefineNameDialog onClose={close} />
    case 'nameManager':
      return <NameManagerDialog onClose={close} />
    case 'createNames':
      return <CreateFromSelectionDialog onClose={close} />
    case 'pivot':
      return <PivotDialog onClose={close} edit={d.props?.edit as { sheetUid: number; id: string } | undefined} />
    case 'subtotal':
      return <SubtotalDialog onClose={close} />
    case 'forecast':
      return <ForecastDialog onClose={close} />
    case 'dataTable':
      return <DataTableDialog onClose={close} />
    case 'scenarioManager':
      return <ScenarioManagerDialog onClose={close} />
    case 'symbol':
      return <SymbolDialog onClose={close} />
    case 'addPen':
      return <AddPenDialog onClose={close} />
    case 'newTableStyle':
      return <NewTableStyleDialog onClose={close} />
    case 'customViews':
      return <CustomViewsDialog onClose={close} />
    case 'editRanges':
      return <EditRangesDialog onClose={close} />
    case 'icons':
      return <IconsDialog onClose={close} />
    case 'evaluateFormula':
      return <EvaluateFormulaDialog onClose={close} />
    case 'spelling':
      return <SpellingDialog onClose={close} />
    case 'sparkline':
      return <SparklineDialog type={d.props?.type as 'line' | 'column' | 'winloss'} onClose={close} />
    case 'quickAnalysis':
      return <QuickAnalysisDialog onClose={close} />
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

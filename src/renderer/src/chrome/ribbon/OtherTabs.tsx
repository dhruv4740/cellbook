// Insert / Page Layout / Formulas / Data / Review / View / Automate / Help tabs.
import { useEffect, useState } from 'react'
import {
  Table20Regular, ImageAdd20Regular, Shapes20Regular, ChartMultiple20Regular, SlideText20Regular,
  TextboxAlignTopLeft20Regular, MathFormula20Regular, Video20Regular, Link20Regular, CommentAdd20Regular,
  SignatureFilled, Sparkle20Regular, DataArea20Regular, PenSparkle20Regular, Calculator20Regular,
  MathSymbols20Regular, ClipboardTextLtr20Regular, ArrowImport20Regular, DataUsage20Regular,
  FlowchartCircle20Regular, ArrowRepeatAll20Regular, TableSettings20Regular, DocumentPageBreak20Regular,
  PanelLeftText20Regular, PageFit20Regular, TextGrammarWand20Regular, TranslateAuto20Regular,
  Note20Regular, CommentMultiple20Regular, LockClosed20Regular, ShieldTask20Regular,
  CodeBlock20Regular, DesktopFlow20Regular, QuestionCircle20Regular, ChatHelp20Regular,
  Mail20Regular, Ruler20Regular, GridDots20Regular, ColumnTriple20Regular, ArrowMoveInward20Regular,
  ZoomIn20Regular, ZoomFit20Regular, TableFreezeRow20Regular, SplitHorizontal20Regular,
  ArrowRepeatAllOff20Regular, AppsListDetail20Regular, DataHistogram20Regular, DataFunnel20Regular,
  Filter20Regular, ArrowSort20Regular, TableSimple20Regular, Print20Regular, Cloud20Regular,
  DataBarVertical20Regular, DataBarHorizontal20Regular, DataTrending20Regular, DataPie20Regular, DataScatter20Regular,
  Eraser20Regular, ArrowUndo20Regular, ArrowRedo20Regular, Add20Regular
} from '@fluentui/react-icons'
import { GenericTab, GGroup, GBtn } from './GenericTab'
import { useStore, store } from '../../model/store'
import { useUi, setUi, getUi, notAvailable } from '../ui'
import { PageSetup } from '../../model/types'
import { PAPER_SIZES, paperOf, marginPresetOf, fitOf } from '../../model/pageSetup'
import { canRemovePageBreak, insertPageBreak, openPageSetup, removePageBreak, resetPageBreaks, setFitPages, setMarginPreset, updatePageSetup } from '../../grid/pageSetupCommands'
import { selectionRef } from '../../dialogs/NameDialogs'
import { FUNCTION_CATEGORIES, RECENTLY_USED } from './functionCategories'
import { MenuItem } from '../../components/Menu'
import * as cmds from '../../grid/commands'
import { COLUMN_BAR_GALLERY, LINE_AREA_GALLERY, PIE_DOUGHNUT_GALLERY, SCATTER_BUBBLE_GALLERY, GalleryEntry } from '../../charts/gallery'
import { insertChartFromSelection } from '../../charts/insertChart'
import { openFile } from '../../io/fileOps'
import { refreshAllPivots } from '../../data/pivot'
import { autoOutline, clearAllOutlines, outlineCommand } from '../../grid/outlineCommands'
import { toggleSheetProtection, toggleWorkbookProtection } from '../../dialogs/ProtectDialogs'

export function ReviewTab() {
  const store = useStore()
  const active = store.active.selection.active
  const key = active.r + ',' + active.c
  const hasComment = !!store.active.layout.comments[key]
  const hasNote = !!store.active.layout.notes[key]
  const noteList = Object.values(store.active.layout.notes)
  const allNotesShown = noteList.length > 0 && noteList.every((n) => n.visible)
  const notesMenu: MenuItem[] = [
    { label: 'New Note', onClick: () => setUi({ noteEditor: { r: active.r, c: active.c } }) },
    { label: 'Edit Note', disabled: !hasNote, onClick: () => setUi({ noteEditor: { r: active.r, c: active.c } }) },
    'separator',
    { label: 'Previous Note', disabled: !noteList.length, onClick: () => cmds.previousNote() },
    { label: 'Next Note', disabled: !noteList.length, onClick: () => cmds.nextNote() },
    { label: 'Show/Hide Note', disabled: !hasNote, checked: !!store.active.layout.notes[key]?.visible, onClick: () => store.updateLayout(store.active, 'Show/Hide Note', (l) => { const n = l.notes[key]; if (n) l.notes[key] = { ...n, visible: !n.visible } }) },
    { label: 'Show All Notes', disabled: !Object.keys(store.active.layout.notes).length, checked: allNotesShown, onClick: () => store.updateLayout(store.active, 'Show All Notes', (l) => { for (const k in l.notes) l.notes[k] = { ...l.notes[k], visible: !allNotesShown } }) },
    'separator',
    { label: 'Convert to Comments', disabled: !Object.keys(store.active.layout.notes).length, onClick: () => store.convertNotesToComments() }
  ]
  const groups: GGroup[] = [
    { label: 'Proofing', big: [
      { label: 'Spelling', icon: <TextGrammarWand20Regular fontSize={28} />, shortcut: 'F7', onClick: () => setUi({ dialog: { kind: 'spelling' } }) },
      { label: 'Thesaurus', icon: <ChatHelp20Regular fontSize={28} /> },
      { label: 'Workbook Statistics', onClick: () => setUi({ dialog: { kind: 'message', props: { text: workbookStats() } } }), icon: <DataUsage20Regular fontSize={28} /> }
    ] },
    { label: 'Accessibility', big: [{ label: 'Check Accessibility', icon: <ShieldTask20Regular fontSize={28} />, onClick: () => setUi({ dialog: { kind: 'accessibility' } }) }] },
    { label: 'Language', big: [{ label: 'Translate', icon: <TranslateAuto20Regular fontSize={28} /> }] },
    { label: 'Changes', big: [{ label: 'Show Changes', icon: <CommentMultiple20Regular fontSize={28} />, disabled: true }] },
    { label: 'Comments', big: [
      { label: 'New Comment', icon: <CommentAdd20Regular fontSize={28} />, onClick: () => setUi({ commentPanel: { r: active.r, c: active.c } }) },
      { label: 'Delete', icon: <CommentMultiple20Regular fontSize={28} />, disabled: !hasComment, onClick: () => { if (hasComment) store.deleteComment(store.active, active.r, active.c) } },
      { label: 'Previous', icon: <CommentMultiple20Regular fontSize={28} />, disabled: !Object.keys(store.active.layout.comments).length, onClick: () => cmds.previousComment() },
      { label: 'Next', icon: <CommentMultiple20Regular fontSize={28} />, disabled: !Object.keys(store.active.layout.comments).length, onClick: () => cmds.nextComment() },
      { label: 'Show Comments', icon: <CommentMultiple20Regular fontSize={28} />, onClick: () => setUi({ commentPanel: { r: active.r, c: active.c } }) }
    ] },
    { label: 'Notes', big: [{ label: 'Notes', icon: <Note20Regular fontSize={28} />, menu: notesMenu }] },
    { label: 'Protect', big: [
      { label: store.active.layout.protection ? 'Unprotect Sheet' : 'Protect Sheet', icon: <LockClosed20Regular fontSize={28} />, wide: true, onClick: () => toggleSheetProtection(store.active) },
      { label: store.bookProtection ? 'Unprotect Workbook' : 'Protect Workbook', icon: <LockClosed20Regular fontSize={28} />, wide: true, onClick: () => toggleWorkbookProtection() },
      { label: 'Allow Edit Ranges', icon: <ShieldTask20Regular fontSize={28} />, disabled: !!store.active.layout.protection, onClick: () => setUi({ dialog: { kind: 'editRanges' } }) },
      { label: 'Unshare Workbook', icon: <ShieldTask20Regular fontSize={28} /> }
    ] },
    { label: 'Ink', big: [{ label: 'Hide Ink', icon: <PenSparkle20Regular fontSize={28} /> }] }
  ]
  return <GenericTab groups={groups} />
}

function workbookStats(): string {
  const sh = store.active
  const used = store.usedRange(sh)
  let cells = 0, formulas = 0
  if (used) for (let r = used.r1; r <= used.r2; r++) for (let c = used.c1; c <= used.c2; c++) { const raw = store.getRaw(sh, r, c); if (raw) { cells++; if (raw.startsWith('=')) formulas++ } }
  const l = sh.layout
  return [
    `Current sheet: ${sh.name}`,
    `  End of sheet: ${used ? cellName(used.r2, used.c2) : 'A1'}`,
    `  Cells with data: ${cells.toLocaleString()}`,
    `  Formulas: ${formulas.toLocaleString()}`,
    `  Notes: ${Object.keys(l.notes).length}   Comments: ${Object.keys(l.comments).length}`,
    `  Charts: ${(l.charts ?? []).length}`,
    '',
    'Workbook',
    `  Sheets: ${store.sheets.length}`,
    `  Names: ${store.names.filter((n) => !/^_xlnm\./i.test(n.name)).length}`
  ].join('\n')
}
function cellName(r: number, c: number): string {
  let s = '', n = c + 1
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) }
  return s + (r + 1)
}

// Excel 2021's Draw tab: pen/highlighter/eraser tools for inking on the sheet. Commands are inert
// (no touch/pen input to drive them in this environment) but the real groups and tools are present.
export function DrawTab() {
  const pen = (label: string, w = 30): GBtn => ({ label, icon: <PenSparkle20Regular fontSize={28} />, noLabel: true, width: w })
  const groups: GGroup[] = [
    { label: 'Input Mode', items: [{ label: 'Draw', icon: <PenSparkle20Regular fontSize={28} />, width: 44 }] },
    { label: 'Undo', items: [[
      { label: 'Undo', icon: <ArrowUndo20Regular />, iconOnly: true, onClick: () => store.undo() },
      { label: 'Redo', icon: <ArrowRedo20Regular />, iconOnly: true, onClick: () => store.redo() }
    ]] },
    { label: 'Drawing Tools', items: [
      { label: 'Select Objects', icon: <DesktopFlow20Regular fontSize={28} />, noLabel: true, width: 40 },
      { label: 'Lasso Select', icon: <DesktopFlow20Regular fontSize={28} />, noLabel: true, width: 40 },
      pen('Eraser'), pen('Pen: Black'), pen('Pen: Red'), pen('Pencil'), pen('Highlighter: Yellow'), pen('Pen: Blue'), pen('Pen: Green'),
      { label: 'Add', icon: <Add20Regular fontSize={28} />, width: 44, menu: [{ label: 'Add Pen...', onClick: () => notAvailable('Add Pen') }] }
    ] },
    { label: 'Convert', items: [
      { label: 'Ink to Shape', icon: <Shapes20Regular fontSize={28} />, width: 50 },
      { label: 'Ink to Math', icon: <MathSymbols20Regular fontSize={28} />, width: 50 }
    ] },
    { label: 'Replay', items: [{ label: 'Ink Replay', icon: <ArrowRepeatAll20Regular fontSize={28} />, width: 48 }] },
    { label: 'Help', items: [{ label: 'Ink Help', icon: <QuestionCircle20Regular fontSize={28} />, width: 44 }] }
  ]
  return <GenericTab groups={groups} />
}

export function HelpTab() {
  const groups: GGroup[] = [
    { label: 'Help', items: [
      { label: 'Help', icon: <QuestionCircle20Regular fontSize={28} />, shortcut: 'F1', width: 40 },
      { label: 'Contact Support', icon: <ChatHelp20Regular fontSize={28} />, width: 56, onClick: () => notAvailable('Contact Support') },
      { label: 'Feedback', icon: <Mail20Regular fontSize={28} />, width: 56, onClick: () => notAvailable('Feedback') },
      { label: 'Show Training', icon: <Sparkle20Regular fontSize={28} />, width: 56 }
    ] },
    { label: 'Community', items: [
      { label: 'Community', icon: <ChatHelp20Regular fontSize={28} />, width: 64 },
      { label: 'Excel Blog', icon: <Note20Regular fontSize={28} />, width: 48 }
    ] },
    { label: 'Mobile', items: [{ label: 'Get Excel Mobile App', icon: <Mail20Regular fontSize={28} />, width: 84 }] }
  ]
  return <GenericTab groups={groups} />
}

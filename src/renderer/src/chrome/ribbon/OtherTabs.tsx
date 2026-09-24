// Insert / Page Layout / Formulas / Data / Review / View / Automate / Help tabs.
import { useState } from 'react'
import {
  Table20Regular, ImageAdd20Regular, Shapes20Regular, ChartMultiple20Regular, SlideText20Regular,
  TextboxAlignTopLeft20Regular, MathFormula20Regular, Video20Regular, Link20Regular, CommentAdd20Regular,
  SignatureFilled, Sparkle20Regular, DataArea20Regular, PenSparkle20Regular, Calculator20Regular,
  MathSymbols20Regular, ClipboardTextLtr20Regular, ArrowImport20Regular, DataUsage20Regular,
  FlowchartCircle20Regular, ArrowRepeatAll20Regular, TableSettings20Regular, DocumentPageBreak20Regular,
  PanelLeftText20Regular, PageFit20Regular, TextGrammarWand20Regular, TranslateAuto20Regular,
  Note20Regular, CommentMultiple20Regular, LockClosed20Regular, ShieldTask20Regular, Play20Regular,
  RecordStop20Regular, CodeBlock20Regular, DesktopFlow20Regular, QuestionCircle20Regular, ChatHelp20Regular,
  Mail20Regular, Ruler20Regular, GridDots20Regular, ColumnTriple20Regular, ArrowMoveInward20Regular,
  ZoomIn20Regular, ZoomFit20Regular, TableFreezeRow20Regular, SplitHorizontal20Regular,
  ArrowRepeatAllOff20Regular, AppsListDetail20Regular, DataHistogram20Regular, DataFunnel20Regular,
  Filter20Regular, ArrowSort20Regular, TableSimple20Regular, Print20Regular, Cloud20Regular
} from '@fluentui/react-icons'
import { GenericTab, GGroup } from './GenericTab'
import { useStore } from '../../model/store'
import { useUi, setUi } from '../ui'
import { MenuItem } from '../../components/Menu'
import * as cmds from '../../grid/commands'

export function InsertTab() {
  const groups: GGroup[] = [
    { label: 'Tables', big: [
      { label: 'PivotTable', icon: <TableSettings20Regular fontSize={28} /> },
      { label: 'Table', icon: <Table20Regular fontSize={28} />, shortcut: 'Ctrl+T' }
    ] },
    { label: 'Illustrations', big: [
      { label: 'Pictures', icon: <ImageAdd20Regular fontSize={28} /> },
      { label: 'Shapes', icon: <Shapes20Regular fontSize={28} /> },
      { label: 'Icons', icon: <Sparkle20Regular fontSize={28} /> },
      { label: 'SmartArt', icon: <FlowchartCircle20Regular fontSize={28} /> }
    ] },
    { label: 'Add-ins', big: [
      { label: 'Get Add-ins', icon: <DataFunnel20Regular fontSize={28} /> },
      { label: 'My Add-ins', icon: <AppsListDetail20Regular fontSize={28} /> }
    ] },
    { label: 'Charts', big: [
      { label: 'Recommended Charts', icon: <ChartMultiple20Regular fontSize={28} /> },
      { label: 'PivotChart', icon: <DataHistogram20Regular fontSize={28} /> }
    ] },
    { label: 'Tours', big: [{ label: '3D Map', icon: <DataArea20Regular fontSize={28} /> }] },
    { label: 'Sparklines', big: [
      { label: 'Line', icon: <ArrowSort20Regular fontSize={28} /> },
      { label: 'Column', icon: <ColumnTriple20Regular fontSize={28} /> },
      { label: 'Win/Loss', icon: <ArrowRepeatAll20Regular fontSize={28} /> }
    ] },
    { label: 'Filters', big: [{ label: 'Slicer', icon: <Filter20Regular fontSize={28} /> }, { label: 'Timeline', icon: <DocumentPageBreak20Regular fontSize={28} /> }] },
    { label: 'Links', big: [{ label: 'Link', icon: <Link20Regular fontSize={28} />, shortcut: 'Ctrl+K' }] },
    { label: 'Comments', big: [{ label: 'Comment', icon: <CommentAdd20Regular fontSize={28} />, onClick: () => {} }] },
    { label: 'Text', big: [
      { label: 'Text Box', icon: <TextboxAlignTopLeft20Regular fontSize={28} /> },
      { label: 'Header & Footer', icon: <PanelLeftText20Regular fontSize={28} /> },
      { label: 'WordArt', icon: <SignatureFilled fontSize={28} /> },
      { label: 'Signature Line', icon: <SignatureFilled fontSize={28} /> },
      { label: 'Object', icon: <MathSymbols20Regular fontSize={28} /> }
    ] },
    { label: 'Symbols', big: [
      { label: 'Equation', icon: <MathFormula20Regular fontSize={28} /> },
      { label: 'Symbol', icon: <MathSymbols20Regular fontSize={28} /> }
    ] },
    { label: 'Media', big: [{ label: 'Video', icon: <Video20Regular fontSize={28} /> }] }
  ]
  return <GenericTab groups={groups} />
}

export function PageLayoutTab() {
  const store = useStore()
  useUi()
  const layout = store.active.layout
  const groups: GGroup[] = [
    { label: 'Themes', big: [{ label: 'Themes', icon: <PenSparkle20Regular fontSize={28} /> }] },
    { label: 'Page Setup', rows: [
      [{ label: 'Margins', icon: <ArrowMoveInward20Regular /> }, { label: 'Orientation', icon: <DocumentPageBreak20Regular /> }, { label: 'Size', icon: <TableSimple20Regular /> }],
      [{ label: 'Print Area', icon: <Print20Regular /> }, { label: 'Breaks', icon: <SplitHorizontal20Regular /> }, { label: 'Background', icon: <ImageAdd20Regular /> }],
      [{ label: 'Print Titles', icon: <TableFreezeRow20Regular /> }]
    ] },
    { label: 'Scale to Fit', rows: [
      [{ label: 'Width', icon: <PageFit20Regular /> }],
      [{ label: 'Height', icon: <PageFit20Regular /> }],
      [{ label: 'Scale', icon: <ZoomFit20Regular /> }]
    ] },
    { label: 'Sheet Options', rows: [
      [{ label: 'Gridlines: View', icon: <GridDots20Regular />, active: layout.showGridlines, onClick: () => store.updateLayout(store.active, 'Gridlines', (l) => (l.showGridlines = !l.showGridlines)) }],
      [{ label: 'Gridlines: Print', icon: <GridDots20Regular /> }],
      [{ label: 'Headings: View', icon: <ColumnTriple20Regular />, active: layout.showHeadings, onClick: () => store.updateLayout(store.active, 'Headings', (l) => (l.showHeadings = !l.showHeadings)) }],
      [{ label: 'Headings: Print', icon: <ColumnTriple20Regular /> }]
    ], launcher: true, onLauncher: () => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } }) },
    { label: 'Arrange', big: [
      { label: 'Bring Forward', icon: <ArrowRepeatAll20Regular fontSize={28} /> },
      { label: 'Send Backward', icon: <ArrowRepeatAllOff20Regular fontSize={28} /> },
      { label: 'Selection Pane', icon: <AppsListDetail20Regular fontSize={28} /> },
      { label: 'Align', icon: <TableFreezeRow20Regular fontSize={28} /> },
      { label: 'Group', icon: <FlowchartCircle20Regular fontSize={28} /> },
      { label: 'Rotate', icon: <ArrowRepeatAll20Regular fontSize={28} /> }
    ] }
  ]
  return <GenericTab groups={groups} />
}

export function FormulasTab() {
  const store = useStore()
  const active = store.active.selection.active
  const insertFn = (name: string) => setUi({ edit: { r: active.r, c: active.c, text: '=' + name + '(', source: 'bar', mode: 'Enter' } })
  const names = (() => { try { return store.hf.getRegisteredFunctionNames() as string[] } catch { return [] } })()
  const catMenu = (): MenuItem[] => (names.length ? names.slice(0, 60).map((n) => ({ label: n, onClick: () => insertFn(n) })) : [{ label: '(none)', disabled: true }])
  const groups: GGroup[] = [
    { label: 'Function Library', rows: undefined, big: [
      { label: 'Insert Function', icon: <MathFormula20Regular fontSize={28} />, onClick: () => insertFn(''), shortcut: 'Shift+F3' },
      { label: 'AutoSum', icon: <MathFormula20Regular fontSize={28} />, menu: [{ label: 'Sum', onClick: () => cmds.autoSum('SUM') }, { label: 'Average', onClick: () => cmds.autoSum('AVERAGE') }, { label: 'Count Numbers', onClick: () => cmds.autoSum('COUNT') }, { label: 'Max', onClick: () => cmds.autoSum('MAX') }, { label: 'Min', onClick: () => cmds.autoSum('MIN') }], onClick: () => cmds.autoSum('SUM') },
      { label: 'Recently Used', icon: <ArrowRepeatAll20Regular fontSize={28} />, menu: catMenu() },
      { label: 'Financial', icon: <Calculator20Regular fontSize={28} />, menu: catMenu() },
      { label: 'Logical', icon: <FlowchartCircle20Regular fontSize={28} />, menu: catMenu() },
      { label: 'Text', icon: <ClipboardTextLtr20Regular fontSize={28} />, menu: catMenu() },
      { label: 'Date & Time', icon: <DocumentPageBreak20Regular fontSize={28} />, menu: catMenu() },
      { label: 'Lookup & Reference', icon: <ArrowImport20Regular fontSize={28} />, menu: catMenu() },
      { label: 'Math & Trig', icon: <MathSymbols20Regular fontSize={28} />, menu: catMenu() },
      { label: 'More Functions', icon: <DataUsage20Regular fontSize={28} />, menu: catMenu() }
    ] },
    { label: 'Defined Names', big: [
      { label: 'Name Manager', icon: <TableSettings20Regular fontSize={28} /> },
      { label: 'Define Name', icon: <TableSettings20Regular fontSize={28} /> },
      { label: 'Use in Formula', icon: <MathFormula20Regular fontSize={28} /> },
      { label: 'Create from Selection', icon: <TableSettings20Regular fontSize={28} /> }
    ] },
    { label: 'Formula Auditing', big: [
      { label: 'Trace Precedents', icon: <ArrowImport20Regular fontSize={28} /> },
      { label: 'Trace Dependents', icon: <ArrowRepeatAll20Regular fontSize={28} /> },
      { label: 'Remove Arrows', icon: <ArrowRepeatAllOff20Regular fontSize={28} /> },
      { label: 'Show Formulas', icon: <ClipboardTextLtr20Regular fontSize={28} />, shortcut: 'Ctrl+`' },
      { label: 'Error Checking', icon: <ShieldTask20Regular fontSize={28} /> },
      { label: 'Watch Window', icon: <AppsListDetail20Regular fontSize={28} /> }
    ] },
    { label: 'Calculation', big: [
      { label: 'Calculation Options', icon: <Calculator20Regular fontSize={28} /> },
      { label: 'Calculate Now', icon: <Calculator20Regular fontSize={28} />, shortcut: 'F9' },
      { label: 'Calculate Sheet', icon: <Calculator20Regular fontSize={28} /> }
    ] }
  ]
  return <GenericTab groups={groups} />
}

export function DataTab() {
  const groups: GGroup[] = [
    { label: 'Get & Transform Data', big: [
      { label: 'Get Data', icon: <ArrowImport20Regular fontSize={28} /> },
      { label: 'From Text/CSV', icon: <ClipboardTextLtr20Regular fontSize={28} /> },
      { label: 'From Web', icon: <Link20Regular fontSize={28} /> },
      { label: 'Recent Sources', icon: <ArrowRepeatAll20Regular fontSize={28} /> },
      { label: 'Existing Connections', icon: <DataArea20Regular fontSize={28} /> }
    ] },
    { label: 'Queries & Connections', big: [
      { label: 'Refresh All', icon: <ArrowRepeatAll20Regular fontSize={28} /> },
      { label: 'Queries & Connections', icon: <DataArea20Regular fontSize={28} /> }
    ] },
    { label: 'Data Types', big: [{ label: 'Stocks', icon: <DataHistogram20Regular fontSize={28} /> }, { label: 'Geography', icon: <DataArea20Regular fontSize={28} /> }] },
    { label: 'Sort & Filter', rows: [
      [{ label: 'A→Z', icon: <ArrowSort20Regular />, onClick: () => cmds.sortSelection('asc') }, { label: 'Z→A', icon: <ArrowSort20Regular />, onClick: () => cmds.sortSelection('desc') }, { label: 'Sort', icon: <ArrowSort20Regular /> }],
      [{ label: 'Filter', icon: <Filter20Regular /> }, { label: 'Clear', icon: <ArrowRepeatAllOff20Regular /> }, { label: 'Reapply', icon: <ArrowRepeatAll20Regular /> }, { label: 'Advanced', icon: <Filter20Regular /> }]
    ] },
    { label: 'Data Tools', big: [
      { label: 'Text to Columns', icon: <ColumnTriple20Regular fontSize={28} /> },
      { label: 'Flash Fill', icon: <Sparkle20Regular fontSize={28} /> },
      { label: 'Remove Duplicates', icon: <TableSimple20Regular fontSize={28} /> },
      { label: 'Data Validation', icon: <ShieldTask20Regular fontSize={28} /> },
      { label: 'Consolidate', icon: <DataArea20Regular fontSize={28} /> },
      { label: 'Relationships', icon: <FlowchartCircle20Regular fontSize={28} /> }
    ] },
    { label: 'Forecast', big: [
      { label: 'What-If Analysis', icon: <DataHistogram20Regular fontSize={28} /> },
      { label: 'Forecast Sheet', icon: <ChartMultiple20Regular fontSize={28} /> }
    ] },
    { label: 'Outline', big: [
      { label: 'Group', icon: <FlowchartCircle20Regular fontSize={28} /> },
      { label: 'Ungroup', icon: <FlowchartCircle20Regular fontSize={28} /> },
      { label: 'Subtotal', icon: <MathFormula20Regular fontSize={28} /> }
    ] }
  ]
  return <GenericTab groups={groups} />
}

export function ReviewTab() {
  const groups: GGroup[] = [
    { label: 'Proofing', big: [
      { label: 'Spelling', icon: <TextGrammarWand20Regular fontSize={28} />, shortcut: 'F7' },
      { label: 'Thesaurus', icon: <ChatHelp20Regular fontSize={28} /> },
      { label: 'Workbook Statistics', icon: <DataUsage20Regular fontSize={28} /> }
    ] },
    { label: 'Accessibility', big: [{ label: 'Check Accessibility', icon: <ShieldTask20Regular fontSize={28} /> }] },
    { label: 'Insights', big: [{ label: 'Smart Lookup', icon: <Sparkle20Regular fontSize={28} /> }] },
    { label: 'Language', big: [{ label: 'Translate', icon: <TranslateAuto20Regular fontSize={28} /> }] },
    { label: 'Comments', big: [
      { label: 'New Comment', icon: <CommentAdd20Regular fontSize={28} /> },
      { label: 'Delete', icon: <CommentMultiple20Regular fontSize={28} /> },
      { label: 'Previous', icon: <CommentMultiple20Regular fontSize={28} /> },
      { label: 'Next', icon: <CommentMultiple20Regular fontSize={28} /> },
      { label: 'Show Comments', icon: <CommentMultiple20Regular fontSize={28} /> }
    ] },
    { label: 'Notes', big: [{ label: 'Notes', icon: <Note20Regular fontSize={28} /> }] },
    { label: 'Protect', big: [
      { label: 'Protect Sheet', icon: <LockClosed20Regular fontSize={28} /> },
      { label: 'Protect Workbook', icon: <LockClosed20Regular fontSize={28} /> },
      { label: 'Allow Edit Ranges', icon: <ShieldTask20Regular fontSize={28} /> },
      { label: 'Unshare Workbook', icon: <ShieldTask20Regular fontSize={28} /> }
    ] },
    { label: 'Ink', big: [{ label: 'Hide Ink', icon: <PenSparkle20Regular fontSize={28} /> }] }
  ]
  return <GenericTab groups={groups} />
}

export function ViewTab() {
  const store = useStore()
  const ui = useUi()
  const layout = store.active.layout
  const [viewMode, setViewMode] = useState<'normal' | 'pageBreak' | 'pageLayout'>('normal')
  const groups: GGroup[] = [
    { label: 'Sheet View', big: [{ label: 'New Sheet View', icon: <TableSettings20Regular fontSize={28} /> }, { label: 'Sheet View', icon: <TableSettings20Regular fontSize={28} /> }] },
    { label: 'Workbook Views', big: [
      { label: 'Normal', icon: <TableSimple20Regular fontSize={28} />, active: viewMode === 'normal', onClick: () => setViewMode('normal') },
      { label: 'Page Break Preview', icon: <DocumentPageBreak20Regular fontSize={28} />, active: viewMode === 'pageBreak', onClick: () => setViewMode('pageBreak') },
      { label: 'Page Layout', icon: <PageFit20Regular fontSize={28} />, active: viewMode === 'pageLayout', onClick: () => setViewMode('pageLayout') },
      { label: 'Custom Views', icon: <AppsListDetail20Regular fontSize={28} /> }
    ] },
    { label: 'Show', rows: [
      [{ label: 'Ruler', icon: <Ruler20Regular /> }, { label: 'Gridlines', icon: <GridDots20Regular />, active: layout.showGridlines, onClick: () => store.updateLayout(store.active, 'Gridlines', (l) => (l.showGridlines = !l.showGridlines)) }],
      [{ label: 'Formula Bar', icon: <ClipboardTextLtr20Regular />, active: !ui.hideFormulaBar, onClick: () => setUi({ hideFormulaBar: !ui.hideFormulaBar }) }],
      [{ label: 'Headings', icon: <ColumnTriple20Regular />, active: layout.showHeadings, onClick: () => store.updateLayout(store.active, 'Headings', (l) => (l.showHeadings = !l.showHeadings)) }]
    ] },
    { label: 'Zoom', big: [
      { label: 'Zoom', icon: <ZoomIn20Regular fontSize={28} />, onClick: () => setUi({ dialog: { kind: 'zoom' } }) },
      { label: '100%', icon: <ZoomFit20Regular fontSize={28} />, onClick: () => store.updateLayout(store.active, 'Zoom', (l) => (l.zoom = 100)) },
      { label: 'Zoom to Selection', icon: <ZoomIn20Regular fontSize={28} /> }
    ] },
    { label: 'Window', rows: [
      [{ label: 'New Window', icon: <AppsListDetail20Regular /> }, { label: 'Arrange All', icon: <AppsListDetail20Regular /> }, { label: 'Freeze Panes', icon: <TableFreezeRow20Regular />, menu: [
        { label: 'Freeze Panes', onClick: () => cmds.freezePanes('panes') },
        { label: 'Freeze Top Row', onClick: () => cmds.freezePanes('topRow') },
        { label: 'Freeze First Column', onClick: () => cmds.freezePanes('firstColumn') },
        { label: 'Unfreeze Panes', onClick: () => cmds.freezePanes('unfreeze') }
      ], onClick: () => cmds.freezePanes('panes') }],
      [{ label: 'Split', icon: <SplitHorizontal20Regular /> }, { label: 'Hide', icon: <ArrowRepeatAllOff20Regular /> }, { label: 'Unhide', icon: <ArrowRepeatAll20Regular /> }],
      [{ label: 'View Side by Side', icon: <AppsListDetail20Regular /> }, { label: 'Synchronous Scrolling', icon: <ArrowSort20Regular /> }, { label: 'Reset Window Position', icon: <ArrowRepeatAllOff20Regular /> }]
    ] },
    { label: 'Macros', big: [{ label: 'Macros', icon: <CodeBlock20Regular fontSize={28} /> }] }
  ]
  return <GenericTab groups={groups} />
}

export function AutomateTab() {
  const groups: GGroup[] = [
    { label: 'Automate', big: [
      { label: 'Record Actions', icon: <RecordStop20Regular fontSize={28} /> },
      { label: 'Automate a Task', icon: <DesktopFlow20Regular fontSize={28} /> },
      { label: 'Run', icon: <Play20Regular fontSize={28} /> },
      { label: 'Editor', icon: <CodeBlock20Regular fontSize={28} /> },
      { label: 'Power Automate', icon: <FlowchartCircle20Regular fontSize={28} /> }
    ] }
  ]
  return <GenericTab groups={groups} />
}

export function HelpTab() {
  const groups: GGroup[] = [
    { label: 'Help', big: [
      { label: 'Help', icon: <QuestionCircle20Regular fontSize={28} />, shortcut: 'F1' },
      { label: 'Show Training', icon: <ChatHelp20Regular fontSize={28} /> },
      { label: "What's New", icon: <Sparkle20Regular fontSize={28} /> },
      { label: 'Suggest a Feature', icon: <Mail20Regular fontSize={28} /> },
      { label: 'Contact Support', icon: <Mail20Regular fontSize={28} /> }
    ] }
  ]
  return <GenericTab groups={groups} />
}

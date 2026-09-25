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
  Eraser20Regular
} from '@fluentui/react-icons'
import { GenericTab, GGroup } from './GenericTab'
import { useStore, store } from '../../model/store'
import { useUi, setUi } from '../ui'
import { MenuItem } from '../../components/Menu'
import * as cmds from '../../grid/commands'
import { COLUMN_BAR_GALLERY, LINE_AREA_GALLERY, PIE_DOUGHNUT_GALLERY, SCATTER_BUBBLE_GALLERY, GalleryEntry } from '../../charts/gallery'
import { insertChartFromSelection } from '../../charts/insertChart'
import { PageSetup, PaperSize, PAPER_SIZES, MARGIN_PRESETS, defaultPageSetup, normRange } from '../../model/types'
import { WorkbookStore, SheetState } from '../../model/store'

function chartGalleryMenu(entries: GalleryEntry[]): MenuItem[] {
  return entries.map((e) => ({ label: e.label, icon: <ChartMultiple20Regular />, onClick: () => insertChartFromSelection(store, e) }))
}

export function InsertTab() {
  const groups: GGroup[] = [
    { label: 'Tables', big: [
      { label: 'PivotTable', icon: <TableSettings20Regular fontSize={28} />, keytip: 'V' },
      { label: 'Table', icon: <Table20Regular fontSize={28} />, shortcut: 'Ctrl+T', keytip: 'T', onClick: () => cmds.openCreateTableDialog() }
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
      { label: 'Recommended Charts', icon: <ChartMultiple20Regular fontSize={28} />, onClick: () => insertChartFromSelection(store, COLUMN_BAR_GALLERY[0]) },
      { label: 'Insert Column or Bar Chart', icon: <DataBarVertical20Regular fontSize={28} />, wide: true, menu: chartGalleryMenu(COLUMN_BAR_GALLERY) },
      { label: 'Insert Line or Area Chart', icon: <DataTrending20Regular fontSize={28} />, wide: true, menu: chartGalleryMenu(LINE_AREA_GALLERY) },
      { label: 'Insert Pie or Doughnut Chart', icon: <DataPie20Regular fontSize={28} />, wide: true, menu: chartGalleryMenu(PIE_DOUGHNUT_GALLERY) },
      { label: 'Insert Scatter (X, Y) or Bubble Chart', icon: <DataScatter20Regular fontSize={28} />, wide: true, menu: chartGalleryMenu(SCATTER_BUBBLE_GALLERY) },
      { label: 'Maps', icon: <DataBarHorizontal20Regular fontSize={28} /> },
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

/** Undoable PageSetup patch, tolerant of a layout saved before pageSetup existed. */
function updatePageSetup(store: WorkbookStore, sheet: SheetState, label: string, fn: (ps: PageSetup) => void): void {
  store.updateLayout(sheet, label, (l) => {
    const ps = l.pageSetup ?? defaultPageSetup()
    fn(ps)
    l.pageSetup = ps
  })
}

export function PageLayoutTab() {
  const store = useStore()
  useUi()
  const sheet = store.active
  const layout = sheet.layout
  const ps = layout.pageSetup ?? defaultPageSetup()

  const marginsMenu: MenuItem[] = [
    { label: 'Normal', onClick: () => updatePageSetup(store, sheet, 'Margins', (p) => (p.margins = { ...MARGIN_PRESETS.normal })) },
    { label: 'Wide', onClick: () => updatePageSetup(store, sheet, 'Margins', (p) => (p.margins = { ...MARGIN_PRESETS.wide })) },
    { label: 'Narrow', onClick: () => updatePageSetup(store, sheet, 'Margins', (p) => (p.margins = { ...MARGIN_PRESETS.narrow })) },
    { label: 'Custom Margins...', onClick: () => setUi({ dialog: { kind: 'pageSetup', props: { tab: 'Margins' } } }) }
  ]
  const orientationMenu: MenuItem[] = [
    { label: 'Portrait', checked: ps.orientation === 'portrait', onClick: () => updatePageSetup(store, sheet, 'Orientation', (p) => (p.orientation = 'portrait')) },
    { label: 'Landscape', checked: ps.orientation === 'landscape', onClick: () => updatePageSetup(store, sheet, 'Orientation', (p) => (p.orientation = 'landscape')) }
  ]
  const sizeMenu: MenuItem[] = [
    ...(Object.entries(PAPER_SIZES) as [PaperSize, typeof PAPER_SIZES[PaperSize]][]).map(([k, v]) => ({
      label: `${v.label} (${v.w}" x ${v.h}")`,
      checked: ps.paperSize === k,
      onClick: () => updatePageSetup(store, sheet, 'Paper Size', (p) => (p.paperSize = k))
    })),
    'separator' as const,
    { label: 'More Paper Sizes...', onClick: () => setUi({ dialog: { kind: 'pageSetup', props: { tab: 'Page' } } }) }
  ]
  const printAreaMenu: MenuItem[] = [
    { label: 'Set Print Area', onClick: () => updatePageSetup(store, sheet, 'Set Print Area', (p) => (p.printArea = sheet.selection.ranges.map(normRange))) },
    { label: 'Clear Print Area', disabled: !ps.printArea.length, onClick: () => updatePageSetup(store, sheet, 'Clear Print Area', (p) => (p.printArea = [])) },
    { label: 'Add to Print Area', disabled: !ps.printArea.length, onClick: () => updatePageSetup(store, sheet, 'Add to Print Area', (p) => (p.printArea = [...p.printArea, ...sheet.selection.ranges.map(normRange)])) }
  ]
  const breaksMenu: MenuItem[] = [
    {
      label: 'Insert Page Break', onClick: () => updatePageSetup(store, sheet, 'Insert Page Break', (p) => {
        const { r, c } = sheet.selection.active
        if (r > 0 && !p.rowBreaks.includes(r)) p.rowBreaks.push(r)
        if (c > 0 && !p.colBreaks.includes(c)) p.colBreaks.push(c)
      })
    },
    {
      label: 'Remove Page Break', onClick: () => updatePageSetup(store, sheet, 'Remove Page Break', (p) => {
        const { r, c } = sheet.selection.active
        p.rowBreaks = p.rowBreaks.filter((x) => x !== r)
        p.colBreaks = p.colBreaks.filter((x) => x !== c)
      })
    },
    { label: 'Reset All Page Breaks', onClick: () => updatePageSetup(store, sheet, 'Reset All Page Breaks', (p) => { p.rowBreaks = []; p.colBreaks = [] }) }
  ]
  const widthMenu: MenuItem[] = [0, 1, 2, 3, 4, 5].map((n) => ({
    label: n === 0 ? 'Automatic' : `${n} Page${n > 1 ? 's' : ''}`,
    checked: ps.fitToPage && ps.fitToWidth === n,
    onClick: () => updatePageSetup(store, sheet, 'Scale to Fit', (p) => { p.fitToPage = true; p.fitToWidth = n })
  }))
  const heightMenu: MenuItem[] = [0, 1, 2, 3, 4, 5].map((n) => ({
    label: n === 0 ? 'Automatic' : `${n} Page${n > 1 ? 's' : ''}`,
    checked: ps.fitToPage && ps.fitToHeight === n,
    onClick: () => updatePageSetup(store, sheet, 'Scale to Fit', (p) => { p.fitToPage = true; p.fitToHeight = n })
  }))
  const scaleMenu: MenuItem[] = [200, 150, 100, 75, 50, 25, 10].map((n) => ({
    label: `${n}%`,
    checked: !ps.fitToPage && ps.scalePercent === n,
    onClick: () => updatePageSetup(store, sheet, 'Scale', (p) => { p.fitToPage = false; p.scalePercent = n })
  }))

  const groups: GGroup[] = [
    { label: 'Themes', big: [{ label: 'Themes', icon: <PenSparkle20Regular fontSize={28} /> }] },
    { label: 'Page Setup', rows: [
      [
        { label: 'Margins', icon: <ArrowMoveInward20Regular />, menu: marginsMenu },
        { label: 'Orientation', icon: <DocumentPageBreak20Regular />, menu: orientationMenu },
        { label: 'Size', icon: <TableSimple20Regular />, menu: sizeMenu }
      ],
      [
        { label: 'Print Area', icon: <Print20Regular />, menu: printAreaMenu },
        { label: 'Breaks', icon: <SplitHorizontal20Regular />, menu: breaksMenu },
        { label: 'Background', icon: <ImageAdd20Regular /> } // not implemented - see final report
      ],
      [{ label: 'Print Titles', icon: <TableFreezeRow20Regular />, onClick: () => setUi({ dialog: { kind: 'pageSetup', props: { tab: 'Sheet' } } }) }]
    ], launcher: true, onLauncher: () => setUi({ dialog: { kind: 'pageSetup', props: { tab: 'Page' } } }) },
    { label: 'Scale to Fit', rows: [
      [{ label: ps.fitToPage ? (ps.fitToWidth === 0 ? 'Width: Automatic' : `Width: ${ps.fitToWidth}`) : 'Width: Automatic', icon: <PageFit20Regular />, menu: widthMenu }],
      [{ label: ps.fitToPage ? (ps.fitToHeight === 0 ? 'Height: Automatic' : `Height: ${ps.fitToHeight}`) : 'Height: Automatic', icon: <PageFit20Regular />, menu: heightMenu }],
      [{ label: `Scale: ${ps.fitToPage ? '(fit to)' : ps.scalePercent + '%'}`, icon: <ZoomFit20Regular />, menu: scaleMenu }]
    ] },
    { label: 'Sheet Options', rows: [
      [{ label: 'Gridlines: View', icon: <GridDots20Regular />, active: layout.showGridlines, onClick: () => store.updateLayout(sheet, 'Gridlines', (l) => (l.showGridlines = !l.showGridlines)) }],
      [{ label: 'Gridlines: Print', icon: <GridDots20Regular />, active: ps.gridlines, onClick: () => updatePageSetup(store, sheet, 'Print Gridlines', (p) => (p.gridlines = !p.gridlines)) }],
      [{ label: 'Headings: View', icon: <ColumnTriple20Regular />, active: layout.showHeadings, onClick: () => store.updateLayout(sheet, 'Headings', (l) => (l.showHeadings = !l.showHeadings)) }],
      [{ label: 'Headings: Print', icon: <ColumnTriple20Regular />, active: ps.headings, onClick: () => updatePageSetup(store, sheet, 'Print Headings', (p) => (p.headings = !p.headings)) }]
    ], launcher: true, onLauncher: () => setUi({ dialog: { kind: 'pageSetup', props: { tab: 'Sheet' } } }) },
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
      { label: 'Insert Function', icon: <MathFormula20Regular fontSize={28} />, onClick: () => insertFn(''), shortcut: 'Shift+F3', keytip: 'FX' },
      { label: 'AutoSum', icon: <MathFormula20Regular fontSize={28} />, menu: [{ label: '&Sum', onClick: () => cmds.autoSum('SUM') }, { label: '&Average', onClick: () => cmds.autoSum('AVERAGE') }, { label: 'Count &Numbers', onClick: () => cmds.autoSum('COUNT') }, { label: 'M&ax', onClick: () => cmds.autoSum('MAX') }, { label: 'M&in', onClick: () => cmds.autoSum('MIN') }], onClick: () => cmds.autoSum('SUM'), keytip: 'U' },
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
  const store = useStore()
  const filtered = !!store.active.layout.autoFilter
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
      [{ label: 'A→Z', icon: <ArrowSort20Regular />, onClick: () => cmds.sortSelection('asc'), keytip: 'SA' }, { label: 'Z→A', icon: <ArrowSort20Regular />, onClick: () => cmds.sortSelection('desc'), keytip: 'SD' }, { label: 'Sort', icon: <ArrowSort20Regular />, onClick: () => cmds.openSortDialog(), keytip: 'S' }],
      [
        { label: 'Filter', icon: <Filter20Regular />, active: filtered, onClick: () => store.toggleAutoFilter(), keytip: 'T' },
        { label: 'Clear', icon: <ArrowRepeatAllOff20Regular />, onClick: () => store.clearAllFilters(), keytip: 'FC' },
        { label: 'Reapply', icon: <ArrowRepeatAll20Regular />, onClick: () => store.reapplyFilter(), keytip: 'Y' },
        { label: 'Advanced', icon: <Filter20Regular />, keytip: 'AD' }
      ]
    ] },
    { label: 'Data Tools', big: [
      { label: 'Text to Columns', icon: <ColumnTriple20Regular fontSize={28} />, onClick: () => setUi({ dialog: { kind: 'textToColumns' } }) },
      { label: 'Flash Fill', icon: <Sparkle20Regular fontSize={28} /> },
      { label: 'Remove Duplicates', icon: <TableSimple20Regular fontSize={28} />, onClick: () => setUi({ dialog: { kind: 'removeDuplicates' } }) },
      { label: 'Data Validation', icon: <ShieldTask20Regular fontSize={28} />, onClick: () => setUi({ dialog: { kind: 'dataValidation' } }) },
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
  const store = useStore()
  const active = store.active.selection.active
  const key = active.r + ',' + active.c
  const hasComment = !!store.active.layout.comments[key]
  const hasNote = !!store.active.layout.notes[key]
  const notesMenu: MenuItem[] = [
    { label: 'New Note', onClick: () => setUi({ noteEditor: { r: active.r, c: active.c } }) },
    { label: 'Edit Note', disabled: !hasNote, onClick: () => setUi({ noteEditor: { r: active.r, c: active.c } }) },
    'separator',
    { label: 'Previous Note', onClick: () => cmds.previousNote() },
    { label: 'Next Note', onClick: () => cmds.nextNote() },
    { label: 'Show/Hide Note', disabled: !hasNote, onClick: () => setUi({ noteEditor: { r: active.r, c: active.c } }) },
    { label: 'Show All Notes', disabled: !Object.keys(store.active.layout.notes).length, onClick: () => cmds.nextNote() },
    'separator',
    { label: 'Convert to Comments', disabled: !Object.keys(store.active.layout.notes).length, onClick: () => store.convertNotesToComments() }
  ]
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
      { label: 'New Comment', icon: <CommentAdd20Regular fontSize={28} />, onClick: () => setUi({ commentPanel: { r: active.r, c: active.c } }) },
      { label: 'Delete', icon: <CommentMultiple20Regular fontSize={28} />, onClick: () => { if (hasComment) store.deleteComment(store.active, active.r, active.c) } },
      { label: 'Previous', icon: <CommentMultiple20Regular fontSize={28} />, onClick: () => cmds.previousComment() },
      { label: 'Next', icon: <CommentMultiple20Regular fontSize={28} />, onClick: () => cmds.nextComment() },
      { label: 'Show Comments', icon: <CommentMultiple20Regular fontSize={28} />, onClick: () => setUi({ commentPanel: { r: active.r, c: active.c } }) }
    ] },
    { label: 'Notes', big: [{ label: 'Notes', icon: <Note20Regular fontSize={28} />, menu: notesMenu, onClick: () => setUi({ noteEditor: { r: active.r, c: active.c } }) }] },
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
  const viewMode = ui.sheetViewMode
  const setViewMode = (m: typeof ui.sheetViewMode) => setUi({ sheetViewMode: m })

  // Other open workbook windows, for the Switch Windows dropdown. Excel keeps one window per
  // workbook, so this doubles as the app's window list.
  const [wins, setWins] = useState<{ id: number; title: string; current: boolean }[]>([])
  useEffect(() => {
    void window.api?.listWindows?.().then((l) => l && setWins(l))
    return window.api?.onWindowListChanged?.((l) => setWins(l))
  }, [])
  const switchWindowsMenu: MenuItem[] = wins.length
    ? wins.map((w) => ({ label: w.title, checked: w.current, onClick: () => void window.api?.activateWindow(w.id) }))
    : [{ label: 'No other windows open', disabled: true }]

  const groups: GGroup[] = [
    { label: 'Sheet View', big: [{ label: 'New Sheet View', icon: <TableSettings20Regular fontSize={28} /> }, { label: 'Sheet View', icon: <TableSettings20Regular fontSize={28} /> }] },
    { label: 'Workbook Views', big: [
      { label: 'Normal', icon: <TableSimple20Regular fontSize={28} />, active: viewMode === 'normal', onClick: () => setViewMode('normal') },
      { label: 'Page Break Preview', icon: <DocumentPageBreak20Regular fontSize={28} />, active: viewMode === 'pageBreak', onClick: () => setViewMode('pageBreak') },
      { label: 'Page Layout', icon: <PageFit20Regular fontSize={28} />, active: viewMode === 'pageLayout', onClick: () => setViewMode('pageLayout') },
      { label: 'Custom Views', icon: <AppsListDetail20Regular fontSize={28} /> }
    ] },
    { label: 'Show', rows: [
      [{ label: 'Ruler', icon: <Ruler20Regular />, keytip: 'VR' }, { label: 'Gridlines', icon: <GridDots20Regular />, active: layout.showGridlines, onClick: () => store.updateLayout(store.active, 'Gridlines', (l) => (l.showGridlines = !l.showGridlines)), keytip: 'VG' }],
      [{ label: 'Formula Bar', icon: <ClipboardTextLtr20Regular />, active: !ui.hideFormulaBar, onClick: () => setUi({ hideFormulaBar: !ui.hideFormulaBar }), keytip: 'VF' }],
      [{ label: 'Headings', icon: <ColumnTriple20Regular />, active: layout.showHeadings, onClick: () => store.updateLayout(store.active, 'Headings', (l) => (l.showHeadings = !l.showHeadings)), keytip: 'VH' }]
    ] },
    { label: 'Zoom', big: [
      { label: 'Zoom', icon: <ZoomIn20Regular fontSize={28} />, onClick: () => setUi({ dialog: { kind: 'zoom' } }) },
      { label: '100%', icon: <ZoomFit20Regular fontSize={28} />, onClick: () => store.updateLayout(store.active, 'Zoom', (l) => (l.zoom = 100)) },
      { label: 'Zoom to Selection', icon: <ZoomIn20Regular fontSize={28} /> }
    ] },
    { label: 'Window', rows: [
      [
        { label: 'New Window', icon: <AppsListDetail20Regular />, onClick: () => void window.api?.newWindow() },
        { label: 'Arrange All', icon: <AppsListDetail20Regular />, onClick: () => void window.api?.arrangeWindowsPrompt() },
        { label: 'Freeze Panes', icon: <TableFreezeRow20Regular />, menu: [
          { label: 'Freeze Panes', onClick: () => cmds.freezePanes('panes') },
          { label: 'Freeze Top Row', onClick: () => cmds.freezePanes('topRow') },
          { label: 'Freeze First Column', onClick: () => cmds.freezePanes('firstColumn') },
          { label: 'Unfreeze Panes', onClick: () => cmds.freezePanes('unfreeze') }
        ], onClick: () => cmds.freezePanes('panes') }
      ],
      [{ label: 'Split', icon: <SplitHorizontal20Regular /> }, { label: 'Hide', icon: <ArrowRepeatAllOff20Regular /> }, { label: 'Unhide', icon: <ArrowRepeatAll20Regular /> }],
      [
        { label: 'View Side by Side', icon: <AppsListDetail20Regular /> },
        { label: 'Synchronous Scrolling', icon: <ArrowSort20Regular /> },
        { label: 'Reset Window Position', icon: <ArrowRepeatAllOff20Regular /> }
      ],
      [{ label: 'Switch Windows', icon: <AppsListDetail20Regular />, menu: switchWindowsMenu }]
    ] },
    { label: 'Macros', big: [{ label: 'Macros', icon: <CodeBlock20Regular fontSize={28} /> }] }
  ]
  return <GenericTab groups={groups} />
}

// Excel 2021's Draw tab: pen/highlighter/eraser tools for inking on the sheet. Commands are inert
// (no touch/pen input to drive them in this environment) but the real groups and tools are present.
export function DrawTab() {
  const groups: GGroup[] = [
    { label: 'Drawing Tools', big: [
      { label: 'Select', icon: <DesktopFlow20Regular fontSize={28} /> },
      { label: 'Lasso Select', icon: <DesktopFlow20Regular fontSize={28} /> },
      { label: 'Eraser', icon: <Eraser20Regular fontSize={28} /> }
    ] },
    { label: 'Pens', big: [
      { label: 'Pen: Black', icon: <PenSparkle20Regular fontSize={28} /> },
      { label: 'Pen: Red', icon: <PenSparkle20Regular fontSize={28} /> },
      { label: 'Pen: Blue', icon: <PenSparkle20Regular fontSize={28} /> },
      { label: 'Pencil', icon: <PenSparkle20Regular fontSize={28} /> },
      { label: 'Highlighter', icon: <PenSparkle20Regular fontSize={28} /> },
      { label: 'Add Pen', icon: <PenSparkle20Regular fontSize={28} /> }
    ] },
    { label: 'Insert', big: [{ label: 'Drawing Canvas', icon: <Shapes20Regular fontSize={28} /> }] },
    { label: 'Convert', big: [
      { label: 'Ink to Shape', icon: <Shapes20Regular fontSize={28} /> },
      { label: 'Ink to Math', icon: <MathSymbols20Regular fontSize={28} /> }
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

// Insert / Page Layout / Formulas / Data / View tabs, laid out group by group like Excel's own ribbon
// (checked against screenshots of real Excel: scripts/excel-oracle/ribbon_shots.py).
import { ReactNode, useEffect, useState } from 'react'
import {
  Table20Regular, ImageAdd20Regular, Shapes20Regular, ChartMultiple20Regular, TextboxAlignTopLeft20Regular, MathFormula20Regular,
  Link20Regular, CommentAdd20Regular, Sparkle20Regular, DataArea20Regular, PenSparkle20Regular, Calculator20Regular, MathSymbols20Regular,
  ClipboardTextLtr20Regular, ArrowImport20Regular, DataUsage20Regular, FlowchartCircle20Regular, ArrowRepeatAll20Regular, TableSettings20Regular,
  DocumentPageBreak20Regular, PanelLeftText20Regular, PageFit20Regular, ShieldTask20Regular, CodeBlock20Regular, Ruler20Regular, GridDots20Regular,
  ColumnTriple20Regular, ArrowMoveInward20Regular, ZoomIn20Regular, ZoomFit20Regular, TableFreezeRow20Regular, SplitHorizontal20Regular,
  ArrowRepeatAllOff20Regular, AppsListDetail20Regular, DataHistogram20Regular, DataFunnel20Regular, Filter20Regular, ArrowSort20Regular,
  TableSimple20Regular, Print20Regular, DataBarVertical20Regular, DataBarHorizontal20Regular, DataTrending20Regular, DataPie20Regular,
  DataScatter20Regular, Camera20Regular, Cube20Regular, Eye20Regular, Add20Regular, Settings20Regular, Layer20Regular, AlignLeft20Regular,
  ArrowRotateClockwise20Regular, Color20Regular, TextFont20Regular, DocumentTable20Regular, Database20Regular,
  ArrowSortDownLines20Regular, ArrowSortUpLines20Regular, EyeOff20Regular, Window20Regular, WindowNew20Regular, ArrowSync20Regular,
  DataWhisker20Regular, DataWaterfall20Regular, Beaker20Regular
} from '@fluentui/react-icons'
import { GenericTab, GGroup, GBtn } from './GenericTab'
import { SmallButton } from './controls'
import { useStore, store } from '../../model/store'
import { useUi, setUi, getUi, notAvailable, showToast } from '../ui'
import { findInvalidCells } from '../../data/dataValidation'
import { traceStep, removeArrows } from '../../data/trace'
import { nextErrorCell } from '../../data/errorCheck'
import { pickPictureFiles } from '../../data/pictures'
import { colName } from '../../model/types'
import { selectionRef } from '../../dialogs/NameDialogs'
import { FUNCTION_CATEGORIES, RECENTLY_USED } from './functionCategories'
import { MenuItem } from '../../components/Menu'
import * as cmds from '../../grid/commands'
import { COLUMN_BAR_GALLERY, LINE_AREA_GALLERY, PIE_DOUGHNUT_GALLERY, SCATTER_BUBBLE_GALLERY, GalleryEntry } from '../../charts/gallery'
import { insertChartFromSelection } from '../../charts/insertChart'
import { openFile } from '../../io/fileOps'
import { refreshAllPivots } from '../../data/pivot'
import { autoOutline, clearAllOutlines, outlineCommand } from '../../grid/outlineCommands'
import { PAPER_SIZES, paperOf, marginPresetOf, fitOf } from '../../model/pageSetup'
import { canRemovePageBreak, insertPageBreak, openPageSetup, removePageBreak, resetPageBreaks, setFitPages, setMarginPreset, setScale, updatePageSetup } from '../../grid/pageSetupCommands'

function errorCheck(): void {
  const sh = store.active
  const hit = nextErrorCell(sh)
  if (!hit) { setUi({ dialog: { kind: 'message', props: { text: 'The error check is complete for the entire sheet.' } } }); return }
  store.setSelection({ ranges: [{ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }], active: { r: hit.r, c: hit.c } })
  showToast(`${colName(hit.c)}${hit.r + 1}: ${hit.kind}`)
}
function circleInvalid(): void {
  const sh = store.active
  const cells = findInvalidCells(store, sh)
  setUi({ validationCircles: cells.length ? { [sh.uid]: cells } : null })
  if (!cells.length) showToast('No cells with invalid data were found.')
}
const big = (I: typeof Table20Regular): ReactNode => <I fontSize={28} />
const na = (label: string): MenuItem[] => [{ label: `${label}...`, onClick: () => notAvailable(label) }]

function chartGalleryMenu(entries: GalleryEntry[]): MenuItem[] {
  return entries.map((e) => ({ label: e.label, icon: <ChartMultiple20Regular />, onClick: () => insertChartFromSelection(store, e) }))
}

function RCheck({ label, checked, disabled, onChange }: { label: string; checked: boolean; disabled?: boolean; onChange?: (v: boolean) => void }) {
  return (
    <label className={'xl-rb-check' + (disabled ? ' disabled' : '')}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange?.(e.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

// ---------------------------------------------------------------------------------------------------------------- Insert
export function InsertTab() {
  const groups: GGroup[] = [
    { label: 'Tables', items: [
      { label: 'PivotTable', icon: big(TableSettings20Regular), width: 58, keytip: 'V', menu: [
        { label: 'From Table/Range...', onClick: () => setUi({ dialog: { kind: 'pivot' } }) },
        { label: 'From External Data Source...', onClick: () => notAvailable('PivotTable from an external data source') },
        { label: 'From Data Model', onClick: () => notAvailable('PivotTable from the Data Model') }
      ], onClick: () => setUi({ dialog: { kind: 'pivot' } }) },
      { label: 'Recommended PivotTables', icon: big(DocumentTable20Regular), width: 104, onClick: () => setUi({ dialog: { kind: 'pivot' } }) },
      { label: 'Table', icon: big(Table20Regular), shortcut: 'Ctrl+T', keytip: 'T', onClick: () => cmds.formatAsTable(0) }
    ] },
    { label: 'Illustrations', items: [
      { label: 'Pictures', icon: big(ImageAdd20Regular), menu: [{ label: 'This Device...', onClick: () => pickPictureFiles() }, { label: 'Stock Images...', onClick: () => notAvailable('Stock Images') }, { label: 'From Online Sources...', onClick: () => notAvailable('Online Pictures') }] },
      [
        { label: 'Shapes', icon: <Shapes20Regular />, menu: na('Shapes') },
        { label: 'Icons', icon: <Sparkle20Regular />, onClick: () => notAvailable('Icons') },
        { label: '3D Models', icon: <Cube20Regular />, menu: na('3D Models') }
      ],
      [
        { label: 'SmartArt', icon: <FlowchartCircle20Regular />, onClick: () => notAvailable('SmartArt') },
        { label: 'Screenshot', icon: <Camera20Regular />, menu: na('Screenshot') }
      ]
    ] },
    { label: 'Charts', launcher: true, onLauncher: () => insertChartFromSelection(store, COLUMN_BAR_GALLERY[0]), items: [
      { label: 'Recommended Charts', icon: big(ChartMultiple20Regular), width: 104, onClick: () => insertChartFromSelection(store, COLUMN_BAR_GALLERY[0]) },
      [
        { label: 'Insert Column or Bar Chart', icon: <DataBarVertical20Regular />, iconOnly: true, menu: chartGalleryMenu(COLUMN_BAR_GALLERY) },
        { label: 'Insert Line or Area Chart', icon: <DataTrending20Regular />, iconOnly: true, menu: chartGalleryMenu(LINE_AREA_GALLERY) },
        { label: 'Insert Pie or Doughnut Chart', icon: <DataPie20Regular />, iconOnly: true, menu: chartGalleryMenu(PIE_DOUGHNUT_GALLERY) }
      ],
      [
        { label: 'Insert Hierarchy Chart', icon: <DataFunnel20Regular />, iconOnly: true, menu: na('Hierarchy charts') },
        { label: 'Insert Statistic Chart', icon: <DataWhisker20Regular />, iconOnly: true, menu: na('Statistic charts') },
        { label: 'Insert Scatter (X, Y) or Bubble Chart', icon: <DataScatter20Regular />, iconOnly: true, menu: chartGalleryMenu(SCATTER_BUBBLE_GALLERY) }
      ],
      [
        { label: 'Insert Waterfall, Funnel, Stock, Surface or Radar Chart', icon: <DataWaterfall20Regular />, iconOnly: true, menu: na('Waterfall, funnel, stock, surface and radar charts') },
        { label: 'Insert Combo Chart', icon: <DataHistogram20Regular />, iconOnly: true, menu: na('Combo charts') }
      ],
      { label: 'Maps', icon: big(DataBarHorizontal20Regular), menu: na('Maps') },
      { label: 'PivotChart', icon: big(DataHistogram20Regular), menu: na('PivotChart') }
    ] },
    { label: 'Sparklines', items: [
      { label: 'Line', icon: big(ArrowSort20Regular), onClick: () => setUi({ dialog: { kind: 'sparkline', props: { type: 'line' } } }) },
      { label: 'Column', icon: big(ColumnTriple20Regular), onClick: () => setUi({ dialog: { kind: 'sparkline', props: { type: 'column' } } }) },
      { label: 'Win/Loss', icon: big(ArrowRepeatAll20Regular), onClick: () => setUi({ dialog: { kind: 'sparkline', props: { type: 'winloss' } } }) }
    ] },
    { label: 'Filters', items: [
      { label: 'Slicer', icon: big(Filter20Regular), onClick: () => notAvailable('Slicers') },
      { label: 'Timeline', icon: big(DocumentPageBreak20Regular), onClick: () => notAvailable('Timelines') }
    ] },
    { label: 'Links', items: [{ label: 'Link', icon: big(Link20Regular), shortcut: 'Ctrl+K', onClick: () => setUi({ dialog: { kind: 'hyperlink' } }) }] },
    { label: 'Comments', items: [{ label: 'Comment', icon: big(CommentAdd20Regular), width: 58, onClick: () => { const a = store.active.selection.active; setUi({ commentPanel: { r: a.r, c: a.c } }) } }] },
    { label: 'Text', items: [{ label: 'Text', icon: big(TextboxAlignTopLeft20Regular), menu: [
      { label: 'Text Box', icon: <TextboxAlignTopLeft20Regular />, onClick: () => notAvailable('Text Box') },
      { label: 'Header & Footer', icon: <PanelLeftText20Regular />, onClick: () => openPageSetup('Header/Footer') },
      { label: 'WordArt', onClick: () => notAvailable('WordArt') },
      { label: 'Signature Line', onClick: () => notAvailable('Signature Line') },
      { label: 'Object', onClick: () => notAvailable('Object') }
    ] }] },
    { label: 'Symbols', items: [{ label: 'Symbols', icon: big(MathSymbols20Regular), menu: [
      { label: 'Equation', icon: <MathFormula20Regular />, onClick: () => notAvailable('Equation') },
      { label: 'Symbol', icon: <MathSymbols20Regular />, onClick: () => notAvailable('Symbol') }
    ] }] }
  ]
  return <GenericTab groups={groups} />
}

// ---------------------------------------------------------------------------------------------------------- Page Layout
export function PageLayoutTab() {
  const store = useStore()
  useUi()
  const layout = store.active.layout
  const ps = layout.pageSetup ?? {}
  const fit = fitOf(ps)
  const hasPrintArea = store.names.some((n) => /^_xlnm\.Print_Area$/i.test(n.name) && n.sheetUid === store.active.uid)
  const pageOptions = (axis: 'w' | 'h'): ReactNode => (
    <label className="xl-rb-field">
      <span>{axis === 'w' ? 'Width:' : 'Height:'}</span>
      <select value={fit.on ? (axis === 'w' ? fit.w : fit.h) : 0} onChange={(e) => setFitPages(axis, +e.target.value)}>
        {[0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => <option key={n} value={n}>{n === 0 ? 'Automatic' : `${n} page${n > 1 ? 's' : ''}`}</option>)}
      </select>
    </label>
  )
  const scaleBox: ReactNode = (
    <label className="xl-rb-field">
      <span>Scale:</span>
      <input type="number" min={10} max={400} disabled={fit.on} value={ps.scale ?? 100} onChange={(e) => setScale(+e.target.value || 100)} />
      <span>%</span>
    </label>
  )
  const themeMenu = (label: string): MenuItem[] => [{ label: `${label}...`, onClick: () => notAvailable(`Workbook ${label.toLowerCase()}`) }]
  const groups: GGroup[] = [
    { label: 'Themes', items: [
      { label: 'Themes', icon: big(Color20Regular), menu: themeMenu('Themes') },
      [
        { label: 'Colors', icon: <Color20Regular />, menu: themeMenu('Colors') },
        { label: 'Fonts', icon: <TextFont20Regular />, menu: themeMenu('Fonts') },
        { label: 'Effects', icon: <Sparkle20Regular />, menu: themeMenu('Effects') }
      ]
    ] },
    { label: 'Page Setup', launcher: true, onLauncher: () => openPageSetup('Page'), items: [
      { label: 'Margins', icon: big(ArrowMoveInward20Regular), width: 54, menu: [
        ...(['normal', 'wide', 'narrow'] as const).map((m) => ({ label: m[0].toUpperCase() + m.slice(1), checked: marginPresetOf(ps) === m, onClick: () => setMarginPreset(m) })),
        { label: 'Custom Margins...', checked: marginPresetOf(ps) === 'custom', onClick: () => openPageSetup('Margins') }
      ] },
      { label: 'Orientation', icon: big(DocumentPageBreak20Regular), width: 66, menu: [
        { label: 'Portrait', checked: !ps.landscape, onClick: () => updatePageSetup({ landscape: false }, 'Orientation') },
        { label: 'Landscape', checked: !!ps.landscape, onClick: () => updatePageSetup({ landscape: true }, 'Orientation') }
      ] },
      { label: 'Size', icon: big(TableSimple20Regular), width: 40, menu: [
        ...PAPER_SIZES.map((z) => ({ label: `${z.name} (${z.wIn} x ${z.hIn} in)`, checked: paperOf(ps).code === z.code, onClick: () => updatePageSetup({ paperSize: z.code }, 'Paper Size') })),
        { label: 'More Paper Sizes...', onClick: () => openPageSetup('Page') }
      ] },
      { label: 'Print Area', icon: big(Print20Regular), width: 48, menu: [
        { label: '&Set Print Area', onClick: () => store.setPrintArea(store.active, selectionRef().replace(/^=/, '')) },
        { label: '&Clear Print Area', disabled: !hasPrintArea, onClick: () => store.setPrintArea(store.active, null) }
      ] },
      { label: 'Breaks', icon: big(SplitHorizontal20Regular), width: 44, menu: [
        { label: 'Insert Page Break', onClick: insertPageBreak },
        { label: 'Remove Page Break', disabled: !canRemovePageBreak(), onClick: removePageBreak },
        { label: 'Reset All Page Breaks', disabled: !(ps.rowBreaks?.length || ps.colBreaks?.length), onClick: resetPageBreaks }
      ] },
      { label: 'Background', icon: big(ImageAdd20Regular), width: 62, onClick: () => notAvailable('Sheet Background') },
      { label: 'Print Titles', icon: big(TableFreezeRow20Regular), width: 50, onClick: () => openPageSetup('Sheet') }
    ] },
    { label: 'Scale to Fit', launcher: true, onLauncher: () => openPageSetup('Page'), items: [[
      { label: 'Width', icon: null, custom: pageOptions('w') },
      { label: 'Height', icon: null, custom: pageOptions('h') },
      { label: 'Scale', icon: null, custom: scaleBox }
    ]] },
    { label: 'Sheet Options', launcher: true, onLauncher: () => openPageSetup('Sheet'), items: [
      [
        { label: 'Gridlines', icon: null, custom: <div className="xl-rb-head">Gridlines</div> },
        { label: 'View', icon: null, custom: <RCheck label="View" checked={layout.showGridlines} onChange={() => store.updateLayout(store.active, 'Gridlines', (l) => (l.showGridlines = !l.showGridlines))} /> },
        { label: 'Print', icon: null, custom: <RCheck label="Print" checked={!!ps.gridlines} onChange={(v) => updatePageSetup({ gridlines: v })} /> }
      ],
      [
        { label: 'Headings', icon: null, custom: <div className="xl-rb-head">Headings</div> },
        { label: 'View', icon: null, custom: <RCheck label="View" checked={layout.showHeadings} onChange={() => store.updateLayout(store.active, 'Headings', (l) => (l.showHeadings = !l.showHeadings))} /> },
        { label: 'Print', icon: null, custom: <RCheck label="Print" checked={!!ps.headings} onChange={(v) => updatePageSetup({ headings: v })} /> }
      ]
    ] },
    { label: 'Arrange', items: [
      { label: 'Bring Forward', icon: big(Layer20Regular), width: 60, disabled: true, menu: na('Bring Forward') },
      { label: 'Send Backward', icon: big(Layer20Regular), width: 60, disabled: true, menu: na('Send Backward') },
      { label: 'Selection Pane', icon: big(AppsListDetail20Regular), width: 60, onClick: () => notAvailable('Selection Pane') },
      { label: 'Align', icon: big(AlignLeft20Regular), width: 40, disabled: true, menu: na('Align') },
      { label: 'Group', icon: big(FlowchartCircle20Regular), width: 44, disabled: true, menu: na('Group') },
      { label: 'Rotate', icon: big(ArrowRotateClockwise20Regular), width: 44, disabled: true, menu: na('Rotate') }
    ] }
  ]
  return <GenericTab groups={groups} />
}

// -------------------------------------------------------------------------------------------------------------- Formulas
export function FormulasTab() {
  const store = useStore()
  const ui = useUi()
  const active = store.active.selection.active
  const insertFn = (name: string) => setUi({ edit: { r: active.r, c: active.c, text: '=' + name + '(', source: 'bar', mode: 'Enter' } })
  const have = (() => { try { return new Set((store.hf.getRegisteredFunctionNames() as string[]).map((n) => n.toUpperCase())) } catch { return new Set<string>() } })()
  const fnMenu = (list: string[]): MenuItem[] => {
    const ok = list.filter((n) => have.has(n))
    return ok.length ? ok.map((n) => ({ label: n, onClick: () => insertFn(n) })) : [{ label: '(none)', disabled: true }]
  }
  const others = new Set(Object.values(FUNCTION_CATEGORIES).flat())
  const moreFunctions = [...have].filter((n) => !others.has(n) && /^[A-Z]/.test(n) && n !== 'ZZROWVISVER').sort()
  const userNames = store.names.filter((n) => !/^_xlnm\./i.test(n.name) && !n.hidden)
  const useInFormula: MenuItem[] = [
    ...userNames.map((n) => ({
      label: n.name,
      onClick: () => {
        const e = getUi().edit
        if (e) setUi({ edit: { ...e, text: e.text + n.name } })
        else setUi({ edit: { r: active.r, c: active.c, text: '=' + n.name, source: 'bar', mode: 'Enter' } })
      }
    } as MenuItem)),
    ...(userNames.length ? ['separator' as const] : []),
    { label: 'Paste Names...', onClick: () => setUi({ dialog: { kind: 'nameManager' } }) }
  ]
  const lib = (label: string, icon: ReactNode, menu: MenuItem[], width?: number): GBtn => ({ label, icon, menu, width })
  const groups: GGroup[] = [
    { label: 'Function Library', items: [
      { label: 'Insert Function', icon: big(MathFormula20Regular), width: 58, onClick: () => insertFn(''), shortcut: 'Shift+F3', keytip: 'FX' },
      { label: 'AutoSum', icon: big(MathFormula20Regular), width: 58, menu: [{ label: '&Sum', onClick: () => cmds.autoSum('SUM') }, { label: '&Average', onClick: () => cmds.autoSum('AVERAGE') }, { label: 'Count &Numbers', onClick: () => cmds.autoSum('COUNT') }, { label: 'M&ax', onClick: () => cmds.autoSum('MAX') }, { label: 'M&in', onClick: () => cmds.autoSum('MIN') }], onClick: () => cmds.autoSum('SUM'), keytip: 'U' },
      lib('Recently Used', big(ArrowRepeatAll20Regular), fnMenu(RECENTLY_USED), 58),
      lib('Financial', big(Calculator20Regular), fnMenu(FUNCTION_CATEGORIES.Financial), 56),
      lib('Logical', big(FlowchartCircle20Regular), fnMenu(FUNCTION_CATEGORIES.Logical), 48),
      lib('Text', big(ClipboardTextLtr20Regular), fnMenu(FUNCTION_CATEGORIES.Text), 40),
      lib('Date & Time', big(DocumentPageBreak20Regular), fnMenu(FUNCTION_CATEGORIES['Date & Time']), 52),
      lib('Lookup & Reference', big(ArrowImport20Regular), fnMenu(FUNCTION_CATEGORIES['Lookup & Reference']), 76),
      lib('Math & Trig', big(MathSymbols20Regular), fnMenu(FUNCTION_CATEGORIES['Math & Trig']), 50),
      lib('More Functions', big(DataUsage20Regular), [
        { label: 'Statistical', submenu: fnMenu(FUNCTION_CATEGORIES.Statistical) },
        { label: 'Information', submenu: fnMenu(FUNCTION_CATEGORIES.Information) },
        { label: 'Other', submenu: fnMenu(moreFunctions).slice(0, 80) }
      ], 76)
    ] },
    { label: 'Defined Names', items: [
      { label: 'Name Manager', icon: big(TableSettings20Regular), width: 58, shortcut: 'Ctrl+F3', onClick: () => setUi({ dialog: { kind: 'nameManager' } }) },
      [
        { label: 'Define Name', icon: <TableSettings20Regular />, menu: [{ label: 'Define Name...', onClick: () => setUi({ dialog: { kind: 'defineName' } }) }, { label: 'Apply Names...', onClick: () => notAvailable('Apply Names') }], onClick: () => setUi({ dialog: { kind: 'defineName' } }) },
        { label: 'Use in Formula', icon: <MathFormula20Regular />, menu: useInFormula },
        { label: 'Create from Selection', icon: <TableSimple20Regular />, shortcut: 'Ctrl+Shift+F3', onClick: () => setUi({ dialog: { kind: 'createNames' } }) }
      ]
    ] },
    { label: 'Formula Auditing', items: [
      [
        { label: 'Trace Precedents', icon: <ArrowImport20Regular />, onClick: () => traceStep('precedents') },
        { label: 'Trace Dependents', icon: <ArrowRepeatAll20Regular />, onClick: () => traceStep('dependents') },
        { label: 'Remove Arrows', icon: <ArrowRepeatAllOff20Regular />, onClick: () => removeArrows() }
      ],
      [
        { label: 'Show Formulas', icon: <ClipboardTextLtr20Regular />, shortcut: 'Ctrl+`', active: ui.formulasVisible, onClick: () => setUi({ formulasVisible: !getUi().formulasVisible }) },
        { label: 'Error Checking', icon: <ShieldTask20Regular />, onClick: () => errorCheck() },
        { label: 'Evaluate Formula', icon: <Beaker20Regular />, onClick: () => setUi({ dialog: { kind: 'evaluateFormula' } }) }
      ],
      { label: 'Watch Window', icon: big(Eye20Regular), width: 52, onClick: () => setUi({ watchOpen: !getUi().watchOpen }) }
    ] },
    { label: 'Calculation', items: [
      { label: 'Calculation Options', icon: big(Calculator20Regular), width: 68, menu: [
        { label: '&Automatic', checked: store.calcMode === 'auto', onClick: () => store.setCalcMode('auto') },
        { label: 'Automatic Except for &Data Tables', checked: store.calcMode === 'autoNoTable', onClick: () => store.setCalcMode('autoNoTable') },
        { label: '&Manual', checked: store.calcMode === 'manual', onClick: () => store.setCalcMode('manual') }
      ] },
      [
        { label: 'Calculate Now', icon: <Calculator20Regular />, shortcut: 'F9', onClick: () => store.recalculate() },
        { label: 'Calculate Sheet', icon: <TableSimple20Regular />, shortcut: 'Shift+F9', onClick: () => store.recalculateSheet() }
      ]
    ] }
  ]
  return <GenericTab groups={groups} />
}

// ------------------------------------------------------------------------------------------------------------------ Data
export function DataTab() {
  const store = useStore()
  const filtered = !!store.active.layout.autoFilter
  const groups: GGroup[] = [
    { label: 'Get & Transform Data', items: [
      { label: 'Get Data', icon: big(Database20Regular), width: 44, menu: na('Get Data') },
      [
        { label: 'From Text/CSV', icon: <ClipboardTextLtr20Regular />, onClick: () => void openFile() },
        { label: 'From Web', icon: <Link20Regular />, onClick: () => notAvailable('From Web') },
        { label: 'From Table/Range', icon: <Table20Regular />, onClick: () => cmds.formatAsTable(0) }
      ],
      [
        { label: 'Recent Sources', icon: <ArrowRepeatAll20Regular />, onClick: () => notAvailable('Recent Sources') },
        { label: 'Existing Connections', icon: <DataArea20Regular />, onClick: () => notAvailable('Existing Connections') }
      ]
    ] },
    { label: 'Queries & Connections', items: [
      { label: 'Refresh All', icon: big(ArrowSync20Regular), width: 44, menu: [
        { label: 'Refresh All', onClick: () => refreshAll() },
        { label: 'Refresh', onClick: () => refreshAll() }
      ], onClick: () => refreshAll() },
      [
        { label: 'Queries & Connections', icon: <DataArea20Regular />, onClick: () => notAvailable('Queries & Connections') },
        { label: 'Properties', icon: <Settings20Regular />, disabled: true },
        { label: 'Workbook Links', icon: <Link20Regular />, disabled: true }
      ]
    ] },
    { label: 'Sort & Filter', items: [
      [
        { label: 'Sort A to Z', icon: <ArrowSortDownLines20Regular />, iconOnly: true, onClick: () => cmds.sortSelection('asc'), keytip: 'SA' },
        { label: 'Sort Z to A', icon: <ArrowSortUpLines20Regular />, iconOnly: true, onClick: () => cmds.sortSelection('desc'), keytip: 'SD' }
      ],
      { label: 'Sort', icon: big(ArrowSort20Regular), width: 40, onClick: () => cmds.openSortDialog(), keytip: 'S' },
      { label: 'Filter', icon: big(Filter20Regular), width: 40, active: filtered, onClick: () => store.toggleAutoFilter(), keytip: 'T' },
      [
        { label: 'Clear', icon: <ArrowRepeatAllOff20Regular />, onClick: () => store.clearAllFilters(), keytip: 'FC', disabled: !filtered && !store.active.layout.advFilter },
        { label: 'Reapply', icon: <ArrowRepeatAll20Regular />, onClick: () => store.reapplyFilter(), keytip: 'Y', disabled: !filtered },
        { label: 'Advanced', icon: <Filter20Regular />, keytip: 'AD', onClick: () => setUi({ dialog: { kind: 'advancedFilter' } }) }
      ]
    ] },
    { label: 'Data Tools', items: [
      { label: 'Text to Columns', icon: big(ColumnTriple20Regular), width: 56, onClick: () => setUi({ dialog: { kind: 'textToColumns' } }), keytip: 'E' },
      [
        { label: 'Flash Fill', icon: <Sparkle20Regular />, iconOnly: true, shortcut: 'Ctrl+E', keytip: 'FF', onClick: () => cmds.flashFillSelection() },
        { label: 'Remove Duplicates', icon: <TableSimple20Regular />, iconOnly: true, onClick: () => setUi({ dialog: { kind: 'removeDuplicates' } }), keytip: 'M' },
        { label: 'Data Validation', icon: <ShieldTask20Regular />, iconOnly: true, menu: [
          { label: 'Data Validation...', onClick: () => setUi({ dialog: { kind: 'dataValidation' } }) },
          { label: 'Circle Invalid Data', onClick: () => circleInvalid() },
          { label: 'Clear Validation Circles', onClick: () => setUi({ validationCircles: null }) }
        ], onClick: () => setUi({ dialog: { kind: 'dataValidation' } }), keytip: 'V' }
      ],
      [
        { label: 'Consolidate', icon: <DataArea20Regular />, iconOnly: true, keytip: 'N', onClick: () => setUi({ dialog: { kind: 'consolidate' } }) },
        { label: 'Relationships', icon: <FlowchartCircle20Regular />, iconOnly: true, onClick: () => notAvailable('Relationships') },
        { label: 'Manage Data Model', icon: <Database20Regular />, iconOnly: true, onClick: () => notAvailable('Data Model') }
      ]
    ] },
    { label: 'Forecast', items: [
      { label: 'What-If Analysis', icon: big(DataHistogram20Regular), width: 56, menu: [
        { label: '&Scenario Manager...', onClick: () => setUi({ dialog: { kind: 'scenarioManager' } }) },
        { label: '&Goal Seek...', onClick: () => setUi({ dialog: { kind: 'goalSeek' } }) },
        { label: 'Data &Table...', onClick: () => setUi({ dialog: { kind: 'dataTable' } }) }
      ] },
      { label: 'Forecast Sheet', icon: big(ChartMultiple20Regular), width: 52, onClick: () => notAvailable('Forecast Sheet') }
    ] },
    { label: 'Outline', launcher: true, onLauncher: () => setUi({ dialog: { kind: 'outlineSettings' } }), items: [
      [
        { label: 'Group', icon: <FlowchartCircle20Regular />, shortcut: 'Alt+Shift+Right', onClick: () => outlineCommand('group'), menu: [
          { label: '&Group...', onClick: () => outlineCommand('group') },
          { label: '&Auto Outline', onClick: () => autoOutline() }
        ] },
        { label: 'Ungroup', icon: <FlowchartCircle20Regular />, shortcut: 'Alt+Shift+Left', onClick: () => outlineCommand('ungroup'), menu: [
          { label: '&Ungroup...', onClick: () => outlineCommand('ungroup') },
          { label: '&Clear Outline', onClick: () => clearAllOutlines() }
        ] },
        { label: 'Subtotal', icon: <MathFormula20Regular />, onClick: () => setUi({ dialog: { kind: 'subtotal' } }) }
      ],
      [
        { label: 'Show Detail', icon: <Add20Regular />, iconOnly: true, onClick: () => outlineCommand('show') },
        { label: 'Hide Detail', icon: <ArrowRepeatAllOff20Regular />, iconOnly: true, onClick: () => outlineCommand('hide') }
      ]
    ] }
  ]
  return <GenericTab groups={groups} />
}

function refreshAll(): void {
  store.recalculate()
  const e = refreshAllPivots(store)
  if (e.length) setUi({ dialog: { kind: 'message', props: { text: e.join('\n') } } })
}

// ------------------------------------------------------------------------------------------------------------------ View
export function ViewTab() {
  const store = useStore()
  const ui = useUi()
  const layout = store.active.layout

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

  const sheetViewCombo: ReactNode = (
    <select className="xl-rb-select" disabled value="Default" style={{ width: 168 }} onChange={() => undefined}>
      <option>Default</option>
    </select>
  )
  const sheetViewRow: ReactNode = (
    <div className="xl-row">
      <SmallButton icon={<EyeOff20Regular />} label="Keep" disabled />
      <SmallButton icon={<Eye20Regular />} label="Exit" disabled />
      <SmallButton icon={<Add20Regular />} label="New" disabled />
      <SmallButton icon={<Settings20Regular />} label="Options" disabled />
    </div>
  )
  const groups: GGroup[] = [
    { label: 'Sheet View', items: [[
      { label: 'Sheet View', icon: null, custom: sheetViewCombo },
      { label: 'Sheet View actions', icon: null, custom: sheetViewRow }
    ]] },
    { label: 'Workbook Views', items: [
      { label: 'Normal', icon: big(TableSimple20Regular), width: 44, active: true },
      { label: 'Page Break Preview', icon: big(DocumentPageBreak20Regular), width: 62, onClick: () => notAvailable('Page Break Preview') },
      [
        { label: 'Page Layout', icon: <PageFit20Regular />, onClick: () => setUi({ backstage: true, backstageTab: 'print' }) },
        { label: 'Custom Views', icon: <AppsListDetail20Regular />, onClick: () => notAvailable('Custom Views') }
      ]
    ] },
    { label: 'Show', items: [
      [
        { label: 'Ruler', icon: null, custom: <RCheck label="Ruler" checked disabled /> },
        { label: 'Gridlines', icon: null, custom: <RCheck label="Gridlines" checked={layout.showGridlines} onChange={() => store.updateLayout(store.active, 'Gridlines', (l) => (l.showGridlines = !l.showGridlines))} /> },
        { label: 'Formula Bar', icon: null, custom: <RCheck label="Formula Bar" checked={!ui.hideFormulaBar} onChange={() => setUi({ hideFormulaBar: !ui.hideFormulaBar })} /> }
      ],
      [
        { label: 'Headings', icon: null, custom: <RCheck label="Headings" checked={layout.showHeadings} onChange={() => store.updateLayout(store.active, 'Headings', (l) => (l.showHeadings = !l.showHeadings))} /> }
      ]
    ] },
    { label: 'Zoom', items: [
      { label: 'Zoom', icon: big(ZoomIn20Regular), width: 44, onClick: () => setUi({ dialog: { kind: 'zoom' } }) },
      { label: '100%', icon: big(ZoomFit20Regular), width: 44, onClick: () => store.updateLayout(store.active, 'Zoom', (l) => (l.zoom = 100)) },
      { label: 'Zoom to Selection', icon: big(ZoomIn20Regular), width: 62, onClick: () => cmds.zoomToSelection() }
    ] },
    { label: 'Window', items: [
      [
        { label: 'New Window', icon: <WindowNew20Regular />, onClick: () => { void window.api?.newViewWindow?.({ data: store.toData(), fileName: store.fileName, filePath: store.filePath }) } },
        { label: 'Arrange All', icon: <Window20Regular />, onClick: () => setUi({ dialog: { kind: 'arrangeWindows' } }) },
        { label: 'Freeze Panes', icon: <TableFreezeRow20Regular />, menu: [
          { label: 'Freeze Panes', onClick: () => cmds.freezePanes('panes') },
          { label: 'Freeze Top Row', onClick: () => cmds.freezePanes('topRow') },
          { label: 'Freeze First Column', onClick: () => cmds.freezePanes('firstColumn') },
          { label: 'Unfreeze Panes', onClick: () => cmds.freezePanes('unfreeze') }
        ] }
      ],
      [
        { label: 'Split', icon: <SplitHorizontal20Regular />, onClick: () => notAvailable('Split') },
        { label: 'Hide', icon: <EyeOff20Regular />, onClick: () => notAvailable('Hide window') },
        { label: 'Unhide', icon: <Eye20Regular />, disabled: true }
      ],
      [
        { label: 'View Side by Side', icon: <AppsListDetail20Regular />, iconOnly: true, disabled: true },
        { label: 'Synchronous Scrolling', icon: <ArrowSort20Regular />, iconOnly: true, disabled: true },
        { label: 'Reset Window Position', icon: <ArrowRepeatAllOff20Regular />, iconOnly: true, disabled: true }
      ],
      { label: 'Switch Windows', icon: big(AppsListDetail20Regular), width: 62, menu: switchWindowsMenu }
    ] },
    { label: 'Macros', items: [{ label: 'Macros', icon: big(CodeBlock20Regular), width: 48, menu: na('Macros') }] }
  ]
  return <GenericTab groups={groups} />
}

void PenSparkle20Regular
void DataBarHorizontal20Regular

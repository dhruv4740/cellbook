// Home tab: Clipboard, Font, Alignment, Number, Styles, Cells, Editing, Add-ins.
import { ReactNode, useRef, useState } from 'react'
import { useStore } from '../../model/store'
import { useUi, setUi, getUi, notAvailable } from '../ui'
import { hideSheet } from '../sheetCommands'
import { CURRENCY_FORMAT, ACCOUNTING_FORMAT, SHORT_DATE_FORMAT, LONG_DATE_FORMAT, CURRENCY_SYMBOL } from '../../engine/locale'
import { CellStyle, normRange } from '../../model/types'
import { MenuItem } from '../../components/Menu'
import { Group, GroupSep, BigButton, SplitButton, StackSplit, SmallButton, DialogLauncher, Combo, KeyTipBadge } from './controls'
import { ChevronDown12Regular } from '@fluentui/react-icons'
import { ColorMenu } from './ColorMenu'
import * as cmds from '../../grid/commands'
import { getSpecialRanges } from '../../dialogs/find'
import { addCfRule, clearCfRules } from '../../format/conditional'
import {
  ClipboardPaste24Regular, Cut20Regular, Copy20Regular, PaintBrush20Regular,
  TextBold20Regular, TextItalic20Regular, TextUnderline20Regular, TextUnderlineDouble20Regular,
  TextStrikethrough20Regular, TextFont20Regular, FontIncrease20Regular, FontDecrease20Regular,
  BorderAll20Regular, PaintBucket20Regular,
  AlignLeft20Regular, AlignRight20Regular, AlignCenterHorizontal20Regular,
  AlignTop20Regular, AlignBottom20Regular, AlignCenterVertical20Regular,
  TextIndentIncrease20Regular, TextIndentDecrease20Regular, TextWrap20Regular,
  TextAlignCenterRotate27020Regular, TextPercent20Regular,
  ArrowSort20Regular, Filter20Regular, Search20Regular,
  Delete20Regular, ArrowExpand20Regular,
  TableInsertRow20Regular, TableSimple20Regular,
  MathFormula20Regular, TableCellsMerge20Regular, DataFunnel20Regular,
  DataHistogram20Regular
} from './icons'

const FONTS = [
  'Calibri', 'Calibri Light', 'Arial', 'Arial Black', 'Bahnschrift', 'Cambria', 'Candara',
  'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Georgia',
  'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'
]
const SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72']

const NUMFMT: [string, string][] = [
  ['General', 'General'],
  ['Number', '0.00'],
  ['Currency', CURRENCY_FORMAT],
  ['Accounting', ACCOUNTING_FORMAT],
  ['Short Date', SHORT_DATE_FORMAT],
  ['Long Date', LONG_DATE_FORMAT],
  ['Time', 'h:mm:ss AM/PM'],
  ['Percentage', '0.00%'],
  ['Fraction', '# ?/?'],
  ['Scientific', '0.00E+00'],
  ['Text', '@']
]

const CELL_STYLES: [string, Partial<CellStyle>][] = [
  ['Normal', {}],
  ['Bad', { color: '#9C0006', fill: '#FFC7CE' }],
  ['Good', { color: '#006100', fill: '#C6EFCE' }],
  ['Neutral', { color: '#9C6500', fill: '#FFEB9C' }],
  ['Calculation', { color: '#FA7D00', fill: '#F2F2F2', bold: true }],
  ['Check Cell', { color: '#FFFFFF', fill: '#A5A5A5', bold: true }],
  ['Explanatory Text', { color: '#7F7F7F', italic: true }],
  ['Input', { color: '#3F3F76', fill: '#FFCC99' }],
  ['Linked Cell', { color: '#FA7D00', underline: 'single' }],
  ['Note', { fill: '#FFFFCC' }],
  ['Output', { color: '#3F3F3F', fill: '#F2F2F2' }],
  ['Warning Text', { color: '#FF0000' }],
  ['Heading 1', { bold: true, fontSize: 15, color: '#1F4E78' }],
  ['Heading 2', { bold: true, fontSize: 13, color: '#1F4E78' }],
  ['Heading 3', { bold: true, fontSize: 11, color: '#1F4E78' }],
  ['Heading 4', { bold: true, fontSize: 11 }],
  ['Title', { bold: true, fontSize: 18 }],
  ['Total', { bold: true }],
  ['20% - Accent1', { fill: '#DAE3F3' }],
  ['40% - Accent1', { fill: '#B4C6E7' }],
  ['60% - Accent1', { fill: '#8EA9DB', color: '#FFFFFF' }],
  ['Accent1', { fill: '#4472C4', color: '#FFFFFF' }],
  ['20% - Accent2', { fill: '#FBE0D0' }],
  ['40% - Accent2', { fill: '#F6C09C' }],
  ['60% - Accent2', { fill: '#F0935E', color: '#FFFFFF' }],
  ['Accent2', { fill: '#ED7D31', color: '#FFFFFF' }],
  ['Comma', { numFmt: '#,##0.00' }],
  ['Currency', { numFmt: CURRENCY_FORMAT }],
  ['Percent', { numFmt: '0%' }]
]

export default function HomeTab() {
  const store = useStore()
  useUi()
  const sheet = store.active
  const active = sheet.selection.active
  const ranges = sheet.selection.ranges
  const style: CellStyle = store.getStyle(sheet, active.r, active.c) ?? {}
  const [colorMenu, setColorMenu] = useState<{ x: number; y: number; mode: 'font' | 'fill' } | null>(null)
  const [lastFontColor, setLastFontColor] = useState('#FF0000')
  const [lastFill, setLastFill] = useState('#FFFF00')
  const [lastBorder, setLastBorder] = useState('bottom')

  const setStyle = (patch: Partial<CellStyle>, label = 'Format Cells') => store.applyStyle(sheet, ranges, patch, label)
  const toggle = (key: 'bold' | 'italic' | 'strike', label: string) => setStyle({ [key]: !style[key] } as Partial<CellStyle>, label)

  const applyBorders = (kind: string) => {
    setLastBorder(kind)
    cmds.applyBorderPreset(kind)
  }


  const decimalAdjust = (delta: 1 | -1) => cmds.adjustDecimal(delta)

  const curFmtName = NUMFMT.find(([, code]) => code === (style.numFmt || 'General'))?.[0] ?? (style.numFmt || 'General')

  const pasteMenu: MenuItem[] = [
    { label: '&Paste', icon: <ClipboardPaste24Regular />, onClick: () => cmds.paste('all') },
    { label: 'Paste &Values', onClick: () => cmds.paste('values') },
    { label: 'Paste &Formulas', onClick: () => cmds.paste('formulas') },
    { label: 'Paste Forma&ts', onClick: () => cmds.paste('formats') },
    { label: 'Transp&ose', onClick: () => cmds.paste('transpose') },
    'separator',
    { label: 'Paste &Link', onClick: () => cmds.paste('link') },
    { label: 'Paste &Special...', onClick: () => setUi({ dialog: { kind: 'pasteSpecial' } }) }
  ]

  const underlineMenu: MenuItem[] = [
    { label: '&Underline', icon: <TextUnderline20Regular />, checked: style.underline === 'single', onClick: () => setStyle({ underline: style.underline === 'single' ? undefined : 'single' }, 'Underline') },
    { label: '&Double Underline', icon: <TextUnderlineDouble20Regular />, checked: style.underline === 'double', onClick: () => setStyle({ underline: style.underline === 'double' ? undefined : 'double' }, 'Underline') }
  ]

  const bordersMenu: MenuItem[] = [
    { label: '&Bottom Border', onClick: () => applyBorders('bottom') },
    { label: '&Top Border', onClick: () => applyBorders('top') },
    { label: '&Left Border', onClick: () => applyBorders('left') },
    { label: '&Right Border', onClick: () => applyBorders('right') },
    { label: '&No Border', onClick: () => applyBorders('none') },
    { label: '&All Borders', onClick: () => applyBorders('all') },
    { label: 'Ou&tside Borders', onClick: () => applyBorders('outside') },
    { label: '&Thick Outside Borders', onClick: () => applyBorders('thickOutside') },
    { label: 'Botto&m Double Border', onClick: () => applyBorders('bottomDouble') },
    { label: 'Thick Botto&m Border', onClick: () => applyBorders('thickBottom') },
    { label: 'T&op and Bottom Border', onClick: () => applyBorders('topBottom') },
    { label: 'Top and Thick Bottom Border', onClick: () => applyBorders('topThickBottom') },
    { label: 'Top and Double Bottom Border', onClick: () => applyBorders('topDoubleBottom') },
    'separator',
    { label: '&More Borders...', onClick: () => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Border' } } }) }
  ]

  const orientationMenu: MenuItem[] = [
    { label: 'Angle Counterclockwise', onClick: () => setStyle({ rotation: 45 }, 'Orientation') },
    { label: 'Angle Clockwise', onClick: () => setStyle({ rotation: -45 }, 'Orientation') },
    { label: 'Vertical Text', onClick: () => setStyle({ rotation: 255 }, 'Orientation') },
    { label: 'Rotate Text Up', onClick: () => setStyle({ rotation: 90 }, 'Orientation') },
    { label: 'Rotate Text Down', onClick: () => setStyle({ rotation: -90 }, 'Orientation') },
    'separator',
    { label: 'Format Cell Alignment...', onClick: () => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Alignment' } } }) }
  ]

  const mergeMenu: MenuItem[] = [
    { label: '&Merge & Center', onClick: () => cmds.mergeCells('center') },
    { label: 'Merge &Across', onClick: () => cmds.mergeCells('across') },
    { label: '&Merge Cells', onClick: () => cmds.mergeCells('merge') },
    { label: '&Unmerge Cells', onClick: () => cmds.mergeCells('unmerge') }
  ]

  const insertMenu: MenuItem[] = [
    { label: '&Insert Cells...', onClick: () => cmds.insertCellsDialog() },
    { label: 'Insert Sheet &Rows', onClick: () => cmds.insertSheetRows() },
    { label: 'Insert Sheet &Columns', onClick: () => cmds.insertSheetColumns() },
    { label: 'Insert &Sheet', onClick: () => store.addSheet() }
  ]
  const deleteMenu: MenuItem[] = [
    { label: '&Delete Cells...', onClick: () => cmds.deleteCellsDialog() },
    { label: 'Delete Sheet &Rows', onClick: () => cmds.deleteSheetRows() },
    { label: 'Delete Sheet &Columns', onClick: () => cmds.deleteSheetColumns() },
    { label: 'Delete &Sheet', onClick: () => store.deleteSheet(store.activeIndex) }
  ]
  // Access keys here match real Excel 2021's Home > Format dropdown exactly (Alt,H,O,<letter>):
  // H row height, A autofit row, W column width, I autofit column, D default width, U hide &
  // unhide (submenu), R rename sheet, M move/copy sheet, T tab color, P protect sheet, L lock
  // cell, E format cells.
  const hideUnhideMenu: MenuItem[] = [
    { label: 'Hide &Rows', onClick: () => cmds.hideRows(true) },
    { label: 'Hide &Columns', onClick: () => cmds.hideColumns(true) },
    { label: 'Hide &Sheet', onClick: () => hideSheet() },
    'separator',
    { label: '&Unhide Rows...', onClick: () => cmds.hideRows(false) },
    { label: 'U&nhide Columns...', onClick: () => cmds.hideColumns(false) },
    { label: 'Unhide S&heet...', disabled: !store.sheets.some((s) => s.layout.sheetHidden === 'hidden'), onClick: () => setUi({ dialog: { kind: 'unhideSheet' } }) }
  ]
  const formatMenu: MenuItem[] = [
    { label: 'Row &Height...', onClick: () => setUi({ dialog: { kind: 'rowHeight' } }) },
    { label: '&AutoFit Row Height', onClick: () => cmds.autoFitRows() },
    { label: 'Column &Width...', onClick: () => setUi({ dialog: { kind: 'columnWidth' } }) },
    { label: 'AutoF&it Column Width', onClick: () => cmds.autoFitColumns() },
    { label: '&Default Width...', onClick: () => setUi({ dialog: { kind: 'columnWidth' } }) },
    'separator',
    { label: 'Hide, &Unhide', submenu: hideUnhideMenu },
    'separator',
    { label: '&Rename Sheet', onClick: () => setUi({ dialog: { kind: 'renameSheet' } }) },
    { label: '&Move or Copy Sheet...', onClick: () => setUi({ dialog: { kind: 'moveCopySheet', props: { index: store.activeIndex } } }) },
    {
      label: 'Ta&b Color',
      submenu: ['#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0', 'No Color'].map((c) => ({
        label: c === 'No Color' ? '&No Color' : c,
        icon: c === 'No Color' ? undefined : <span style={{ display: 'inline-block', width: 12, height: 12, background: c }} />,
        onClick: () => store.updateLayout(store.active, 'Tab Color', (l) => (l.tabColor = c === 'No Color' ? undefined : c))
      }))
    },
    'separator',
    { label: '&Protect Sheet...', onClick: () => notAvailable('Protect Sheet') },
    { label: '&Lock Cell', onClick: () => setStyle({ locked: !style.locked }, 'Lock Cell') },
    { label: 'Format C&ells...', shortcut: 'Ctrl+1', onClick: () => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } }) }
  ]

  const autoSumMenu: MenuItem[] = [
    { label: '&Sum', icon: <MathFormula20Regular />, onClick: () => cmds.autoSum('SUM') },
    { label: '&Average', onClick: () => cmds.autoSum('AVERAGE') },
    { label: 'Count &Numbers', onClick: () => cmds.autoSum('COUNT') },
    { label: '&Max', onClick: () => cmds.autoSum('MAX') },
    { label: 'M&in', onClick: () => cmds.autoSum('MIN') },
    { label: 'More &Functions...', onClick: () => setUi({ edit: { r: active.r, c: active.c, text: '=', source: 'bar', mode: 'Enter' } }) }
  ]
  const fillDirMenu: MenuItem[] = [
    { label: '&Down', shortcut: 'Ctrl+D', onClick: () => cmds.fill('down') },
    { label: '&Right', shortcut: 'Ctrl+R', onClick: () => cmds.fill('right') },
    { label: '&Up', onClick: () => cmds.fill('up') },
    { label: '&Left', onClick: () => cmds.fill('left') },
    'separator',
    { label: '&Series...', onClick: () => setUi({ dialog: { kind: 'series' } }) },
    { label: '&Justify', onClick: () => notAvailable('Justify') }
  ]
  const clearMenu: MenuItem[] = [
    { label: 'Clear &All', onClick: () => cmds.clearSelection('all') },
    { label: 'Clear &Formats', onClick: () => cmds.clearSelection('formats') },
    { label: 'Clear &Contents', shortcut: 'Del', onClick: () => cmds.clearSelection('contents') },
    { label: 'Clear Co&mments and Notes', onClick: () => cmds.clearSelection('notes') },
    { label: 'Clear H&yperlinks', onClick: () => cmds.clearSelection('hyperlinks') }
  ]
  const sortMenu: MenuItem[] = [
    { label: 'Sort &A to Z', onClick: () => cmds.sortSelection('asc') },
    { label: 'Sort &Z to A', onClick: () => cmds.sortSelection('desc') },
    { label: 'Custom Sort...', onClick: () => cmds.openSortDialog() },
    'separator',
    { label: '&Filter', icon: <Filter20Regular />, checked: !!store.active.layout.autoFilter, shortcut: 'Ctrl+Shift+L', onClick: () => store.toggleAutoFilter() },
    { label: 'Clear', onClick: () => store.clearAllFilters(), disabled: !store.active.layout.autoFilter },
    { label: 'Reapply', onClick: () => store.reapplyFilter(), disabled: !store.active.layout.autoFilter }
  ]
  const findMenu: MenuItem[] = [
    { label: '&Find...', shortcut: 'Ctrl+F', onClick: () => setUi({ dialog: { kind: 'find' } }) },
    { label: '&Replace...', shortcut: 'Ctrl+H', onClick: () => setUi({ dialog: { kind: 'replace' } }) },
    { label: 'Go &To...', shortcut: 'Ctrl+G', onClick: () => setUi({ dialog: { kind: 'goto' } }) },
    { label: 'Go To &Special...', onClick: () => setUi({ dialog: { kind: 'gotoSpecial' } }) },
    'separator',
    {
      label: 'Formulas',
      onClick: () => {
        const ranges = getSpecialRanges(store, 'formulas')
        if (ranges.length > 0) store.setSelection({ ranges, active: { r: ranges[0].r1, c: ranges[0].c1 } })
      }
    },
    {
      label: 'Notes',
      onClick: () => {
        const ranges = getSpecialRanges(store, 'comments')
        if (ranges.length > 0) store.setSelection({ ranges, active: { r: ranges[0].r1, c: ranges[0].c1 } })
      }
    },
    {
      label: 'Constants',
      onClick: () => {
        const ranges = getSpecialRanges(store, 'constants')
        if (ranges.length > 0) store.setSelection({ ranges, active: { r: ranges[0].r1, c: ranges[0].c1 } })
      }
    },
    'separator',
    { label: 'Select Objects', onClick: () => notAvailable('Select Objects') }
  ]

  const cfMenu: MenuItem[] = [
    {
      label: 'Highlight Cells Rules',
      submenu: [
        { label: 'Greater Than...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'greaterThan' } } }) },
        { label: 'Less Than...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'lessThan' } } }) },
        { label: 'Between...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'between' } } }) },
        { label: 'Equal To...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'equal' } } }) },
        { label: 'Text that Contains...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'textContains' } } }) },
        { label: 'A Date Occurring...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'dateOccurring' } } }) },
        { label: 'Duplicate Values...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'duplicateValues' } } }) },
        'separator',
        { label: 'More Rules...', onClick: () => setUi({ dialog: { kind: 'cfNewRule' } }) }
      ]
    },
    {
      label: 'Top/Bottom Rules',
      submenu: [
        { label: 'Top 10 Items...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'top10Items' } } }) },
        { label: 'Top 10%...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'top10Percent' } } }) },
        { label: 'Bottom 10 Items...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'bottom10Items' } } }) },
        { label: 'Bottom 10%...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'bottom10Percent' } } }) },
        { label: 'Above Average...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'aboveAverage' } } }) },
        { label: 'Below Average...', onClick: () => setUi({ dialog: { kind: 'cfQuick', props: { kind: 'belowAverage' } } }) },
        'separator',
        { label: 'More Rules...', onClick: () => setUi({ dialog: { kind: 'cfNewRule' } }) }
      ]
    },
    {
      label: 'Data Bars',
      submenu: [
        { label: 'Blue Data Bar', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#638EC6', gradient: true } }) },
        { label: 'Green Data Bar', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#63C384', gradient: true } }) },
        { label: 'Red Data Bar', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#FF555A', gradient: true } }) },
        { label: 'Orange Data Bar', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#FFB620', gradient: true } }) },
        { label: 'Light Blue Data Bar', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#33CCCC', gradient: true } }) },
        { label: 'Purple Data Bar', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#800080', gradient: true } }) },
        'separator',
        { label: 'Blue Data Bar (Solid)', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#0070C0', gradient: false } }) },
        { label: 'Green Data Bar (Solid)', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#00B050', gradient: false } }) },
        { label: 'Red Data Bar (Solid)', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#C00000', gradient: false } }) },
        { label: 'Orange Data Bar (Solid)', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#FFC000', gradient: false } }) },
        { label: 'Light Blue Data Bar (Solid)', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#00B0F0', gradient: false } }) },
        { label: 'Purple Data Bar (Solid)', onClick: () => addCfRule(store, sheet, { ranges, type: 'dataBar', dataBar: { color: '#7030A0', gradient: false } }) },
        'separator',
        { label: 'More Rules...', onClick: () => setUi({ dialog: { kind: 'cfNewRule' } }) }
      ]
    },
    {
      label: 'Color Scales',
      submenu: [
        { label: 'Green - Yellow - Red Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#F8696B', midColor: '#FFEB84', maxColor: '#63BE7B' } }) },
        { label: 'Red - Yellow - Green Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#63BE7B', midColor: '#FFEB84', maxColor: '#F8696B' } }) },
        { label: 'Green - White - Red Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#F8696B', midColor: '#FFFFFF', maxColor: '#63BE7B' } }) },
        { label: 'Red - White - Green Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#63BE7B', midColor: '#FFFFFF', maxColor: '#F8696B' } }) },
        { label: 'Blue - White - Red Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#F8696B', midColor: '#FFFFFF', maxColor: '#5A8AC6' } }) },
        { label: 'Red - White - Blue Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#5A8AC6', midColor: '#FFFFFF', maxColor: '#F8696B' } }) },
        { label: 'White - Red Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#FFFFFF', maxColor: '#F8696B' } }) },
        { label: 'Red - White Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#F8696B', maxColor: '#FFFFFF' } }) },
        { label: 'Green - White Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#FFFFFF', maxColor: '#63BE7B' } }) },
        { label: 'White - Green Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#63BE7B', maxColor: '#FFFFFF' } }) },
        { label: 'Green - Yellow Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#FFEB84', maxColor: '#63BE7B' } }) },
        { label: 'Yellow - Green Color Scale', onClick: () => addCfRule(store, sheet, { ranges, type: 'colorScale', colorScale: { minColor: '#63BE7B', maxColor: '#FFEB84' } }) },
        'separator',
        { label: 'More Rules...', onClick: () => setUi({ dialog: { kind: 'cfNewRule' } }) }
      ]
    },
    {
      label: 'Icon Sets',
      submenu: [
        { label: '3 Traffic Lights (Unrimmed)', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '3TrafficLights1' } }) },
        { label: '3 Traffic Lights (Rimmed)', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '3TrafficLights2' } }) },
        { label: '3 Arrows (Colored)', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '3Arrows' } }) },
        { label: '3 Symbols (Circled)', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '3Symbols' } }) },
        { label: '3 Flags', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '3Flags' } }) },
        { label: '4 Ratings', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '4Rating' } }) },
        { label: '5 Ratings', onClick: () => addCfRule(store, sheet, { ranges, type: 'iconSet', iconSet: { name: '5Rating' } }) },
        'separator',
        { label: 'More Rules...', onClick: () => setUi({ dialog: { kind: 'cfNewRule' } }) }
      ]
    },
    'separator',
    { label: 'New Rule...', onClick: () => setUi({ dialog: { kind: 'cfNewRule' } }) },
    {
      label: 'Clear Rules',
      submenu: [
        { label: 'Clear Rules from Selected Cells', onClick: () => clearCfRules(store, sheet, 'selection') },
        { label: 'Clear Rules from Entire Sheet', onClick: () => clearCfRules(store, sheet, 'sheet') }
      ]
    },
    { label: 'Manage Rules...', onClick: () => setUi({ dialog: { kind: 'cfManage' } }) }
  ]

  return (
    <>
      <Group label="Clipboard" icon={<ClipboardPaste24Regular fontSize={20} />}>
        <div className="xl-col">
          <SplitButton icon={<ClipboardPaste24Regular fontSize={28} />} label="Paste" tip={{ title: 'Paste (Ctrl+V)', desc: 'Paste the contents of the Clipboard.' }} onClick={() => cmds.paste('all')} menu={pasteMenu} keytip="V" />
        </div>
        <div className="xl-col xl-stack">
          <SmallButton icon={<Cut20Regular />} label="Cut" tip={{ title: 'Cut (Ctrl+X)' }} onClick={() => cmds.cut()} keytip="X" />
          <SmallButton icon={<Copy20Regular />} label="Copy" menu={[{ label: '&Copy', onClick: () => cmds.copy() }, { label: 'Copy as &Picture...', onClick: () => notAvailable('Copy as Picture') }]} tip={{ title: 'Copy (Ctrl+C)' }} onClick={() => cmds.copy()} keytip="C" />
          <SmallButton
            icon={<PaintBrush20Regular />}
            label="Format Painter"
            tip={{ title: 'Format Painter', desc: 'Copy formatting to another selection. Double-click to keep it on.' }}
            active={!!getUi().formatPainter}
            onClick={() => setUi({ formatPainter: getUi().formatPainter ? null : { once: true } })}
            onDoubleClick={() => setUi({ formatPainter: { once: false } })}
            keytip="FP"
          />
        </div>
      </Group>
      <GroupSep />

      <Group label="Font" icon={<TextFont20Regular />}>
        <div className="xl-col">
          <div className="xl-row">
            <Combo value={style.fontFamily ?? 'Calibri'} options={FONTS} width={120} tip={{ title: 'Font' }} onChange={(v) => setStyle({ fontFamily: v }, 'Font')} keytip="FF" />
            <Combo value={String(style.fontSize ?? 11)} options={SIZES} width={48} tip={{ title: 'Font Size' }} onChange={(v) => { const n = +v; if (n > 0) setStyle({ fontSize: n }, 'Font Size') }} keytip="FS" />
            <SmallButton icon={<FontIncrease20Regular />} tip={{ title: 'Increase Font Size' }} onClick={() => setStyle({ fontSize: (style.fontSize ?? 11) + 1 }, 'Font Size')} keytip="FG" />
            <SmallButton icon={<FontDecrease20Regular />} tip={{ title: 'Decrease Font Size' }} onClick={() => setStyle({ fontSize: Math.max(1, (style.fontSize ?? 11) - 1) }, 'Font Size')} keytip="FK" />
          </div>
          <div className="xl-row">
            <SmallButton icon={<TextBold20Regular />} active={!!style.bold} tip={{ title: 'Bold (Ctrl+B)' }} onClick={() => toggle('bold', 'Bold')} keytip="1" />
            <SmallButton icon={<TextItalic20Regular />} active={!!style.italic} tip={{ title: 'Italic (Ctrl+I)' }} onClick={() => toggle('italic', 'Italic')} keytip="2" />
            <SmallButton icon={<TextUnderline20Regular />} active={!!style.underline} tip={{ title: 'Underline (Ctrl+U)' }} split menu={underlineMenu} onClick={() => setStyle({ underline: style.underline ? undefined : 'single' }, 'Underline')} keytip="3" />
            <SmallButton icon={<TextStrikethrough20Regular />} active={!!style.strike} tip={{ title: 'Strikethrough' }} onClick={() => toggle('strike', 'Strikethrough')} keytip="4" />
            <SmallButton icon={<BorderAll20Regular />} tip={{ title: 'Borders' }} split menu={bordersMenu} onClick={() => applyBorders(lastBorder)} keytip="B" />
            <ColorSplit
              icon={<PaintBucket20Regular />}
              color={lastFill}
              tip={{ title: 'Fill Color' }}
              onClick={() => setStyle({ fill: lastFill }, 'Fill Color')}
              onOpen={(x, y) => setColorMenu({ x, y, mode: 'fill' })}
              keytip="H"
            />
            <ColorSplit
              icon={<TextFont20Regular />}
              color={style.color ?? lastFontColor}
              tip={{ title: 'Font Color' }}
              onClick={() => setStyle({ color: lastFontColor }, 'Font Color')}
              onOpen={(x, y) => setColorMenu({ x, y, mode: 'font' })}
              keytip="FC"
            />
          </div>
        </div>
        <DialogLauncher tip={{ title: 'Format Cells: Font', shortcut: 'Ctrl+Shift+F' }} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Font' } } })} />
      </Group>
      <GroupSep />

      <Group label="Alignment" icon={<AlignCenterHorizontal20Regular />}>
        <div className="xl-col">
          <div className="xl-row">
            <SmallButton icon={<AlignTop20Regular />} active={style.vAlign === 'top'} tip={{ title: 'Top Align' }} onClick={() => setStyle({ vAlign: 'top' }, 'Align')} keytip="AT" />
            <SmallButton icon={<AlignCenterVertical20Regular />} active={style.vAlign === 'middle'} tip={{ title: 'Middle Align' }} onClick={() => setStyle({ vAlign: 'middle' }, 'Align')} keytip="AM" />
            <SmallButton icon={<AlignBottom20Regular />} active={style.vAlign === 'bottom'} tip={{ title: 'Bottom Align' }} onClick={() => setStyle({ vAlign: 'bottom' }, 'Align')} keytip="AB" />
            <SmallButton icon={<TextAlignCenterRotate27020Regular />} tip={{ title: 'Orientation' }} menu={orientationMenu} keytip="TQ" />
          </div>
          <div className="xl-row">
            <SmallButton icon={<AlignLeft20Regular />} active={style.hAlign === 'left'} tip={{ title: 'Align Left' }} onClick={() => setStyle({ hAlign: 'left' }, 'Align')} keytip="AL" />
            <SmallButton icon={<AlignCenterHorizontal20Regular />} active={style.hAlign === 'center'} tip={{ title: 'Center' }} onClick={() => setStyle({ hAlign: 'center' }, 'Align')} keytip="AC" />
            <SmallButton icon={<AlignRight20Regular />} active={style.hAlign === 'right'} tip={{ title: 'Align Right' }} onClick={() => setStyle({ hAlign: 'right' }, 'Align')} keytip="AR" />
            <SmallButton icon={<TextWrap20Regular />} active={!!style.wrap} tip={{ title: 'Wrap Text' }} onClick={() => setStyle({ wrap: !style.wrap }, 'Wrap Text')} keytip="W" />
          </div>
          <div className="xl-row">
            <SmallButton icon={<TextIndentDecrease20Regular />} tip={{ title: 'Decrease Indent' }} onClick={() => setStyle({ indent: Math.max(0, (style.indent ?? 0) - 1) }, 'Indent')} keytip="AO" />
            <SmallButton icon={<TextIndentIncrease20Regular />} tip={{ title: 'Increase Indent' }} onClick={() => setStyle({ indent: (style.indent ?? 0) + 1 }, 'Indent')} keytip="AI" />
          </div>
        </div>
        <div className="xl-col xl-align-side">
          <SmallButton icon={<TextWrap20Regular />} label="Wrap Text" active={!!style.wrap} tip={{ title: 'Wrap Text' }} onClick={() => setStyle({ wrap: !style.wrap }, 'Wrap Text')} />
          <SmallButton icon={<TableCellsMerge20Regular />} label="Merge & Center" tip={{ title: 'Merge & Center' }} menu={mergeMenu} onClick={() => cmds.mergeCells('center')} keytip="M" />
        </div>
        <DialogLauncher tip={{ title: 'Format Cells: Alignment' }} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Alignment' } } })} />
      </Group>
      <GroupSep />

      <Group label="Number" icon={<TextPercent20Regular />}>
        <div className="xl-col">
          <Combo value={curFmtName} options={NUMFMT.map(([n]) => n)} width={122} editable={false} tip={{ title: 'Number Format' }} onPick={(name) => { const f = NUMFMT.find(([n]) => n === name); if (f) setStyle({ numFmt: f[1] }, 'Number Format') }} keytip="N" />
          <div className="xl-row">
            <SmallButton icon={<Glyph text={CURRENCY_SYMBOL.length === 1 ? CURRENCY_SYMBOL : '$'} />} tip={{ title: 'Accounting Number Format' }} split menu={[
              { label: '₹ English (India)', onClick: () => setStyle({ numFmt: '_ ₹* #,##0.00_ ;_ ₹* -#,##0.00_ ;_ ₹* "-"??_ ;_ @_ ' }, 'Number Format') },
              { label: '$ English (United States)', onClick: () => setStyle({ numFmt: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' }, 'Number Format') },
              { label: '£ English (United Kingdom)', onClick: () => setStyle({ numFmt: '_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-' }, 'Number Format') },
              { label: '€ Euro (€ 123)', onClick: () => setStyle({ numFmt: '_-[$€-x-euro2] * #,##0.00_-;-[$€-x-euro2] * #,##0.00_-;_-[$€-x-euro2] * "-"??_-;_-@_-' }, 'Number Format') },
              'separator',
              { label: '&More Accounting Formats...', onClick: () => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } }) }
            ]} onClick={() => setStyle({ numFmt: NUMFMT[3][1] }, 'Number Format')} keytip="AN" />
            <SmallButton icon={<Glyph text="%" />} tip={{ title: 'Percent Style (Ctrl+Shift+%)' }} onClick={() => setStyle({ numFmt: '0%' }, 'Number Format')} keytip="P" />
            <SmallButton icon={<Glyph text="9" sub="," />} tip={{ title: 'Comma Style' }} onClick={() => setStyle({ numFmt: '#,##0.00' }, 'Number Format')} keytip="K" />
            <SmallButton icon={<DecimalGlyph more />} tip={{ title: 'Increase Decimal' }} onClick={() => decimalAdjust(1)} keytip="0" />
            <SmallButton icon={<DecimalGlyph />} tip={{ title: 'Decrease Decimal' }} onClick={() => decimalAdjust(-1)} keytip="9" />
          </div>
        </div>
        <DialogLauncher tip={{ title: 'Format Cells: Number' }} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } })} />
      </Group>
      <GroupSep />

      <Group label="Styles" icon={<TableCellsMerge20Regular />}>
        <BigButton wide icon={<TableCellsMerge20Regular fontSize={28} />} label="Conditional Formatting" tip={{ title: 'Conditional Formatting' }} menu={cfMenu} keytip="L" />
        <BigButton wide icon={<TableSimple20Regular fontSize={28} />} label="Format as Table" tip={{ title: 'Format as Table' }} menu={[...cmds.TABLE_STYLES.map((t, i) => ({ label: t.name, icon: <span style={{ display: 'inline-block', width: 14, height: 12, background: `linear-gradient(${t.header} 0 34%, ${t.band} 34% 67%, #fff 67%)`, border: `1px solid ${t.line}` }} />, onClick: () => cmds.formatAsTable(i) } as MenuItem)), 'separator', { label: 'New Table Style...', onClick: () => notAvailable('New Table Style') }]} keytip="T" />
        <BigButton
          wide
          icon={<TableInsertRow20Regular fontSize={28} />}
          label="Cell Styles"
          tip={{ title: 'Cell Styles' }}
          menu={CELL_STYLES.map(([name, s]) => ({ label: name, onClick: () => setStyle(s, 'Cell Styles') } as MenuItem))}
          keytip="J"
        />
      </Group>
      <GroupSep />

      <Group label="Cells" icon={<TableInsertRow20Regular />}>
        <div className="xl-col xl-stack">
          <StackSplit icon={<TableInsertRow20Regular />} label="Insert" tip={{ title: 'Insert Cells' }} menu={insertMenu} onClick={() => cmds.insertCellsDialog()} keytip="I" />
          <StackSplit icon={<Delete20Regular />} label="Delete" tip={{ title: 'Delete Cells' }} menu={deleteMenu} onClick={() => cmds.deleteCellsDialog()} keytip="D" />
          <StackSplit icon={<ArrowExpand20Regular />} label="Format" tip={{ title: 'Format' }} menu={formatMenu} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } })} keytip="O" />
        </div>
      </Group>
      <GroupSep />

      <Group label="Editing" icon={<ArrowSort20Regular />}>
        <div className="xl-col xl-stack">
          <StackSplit icon={<MathFormula20Regular />} label="AutoSum" tip={{ title: 'AutoSum (Alt+=)' }} menu={autoSumMenu} onClick={() => cmds.autoSum('SUM')} keytip="U" />
          <StackSplit icon={<TableInsertRow20Regular />} label="Fill" tip={{ title: 'Fill' }} menu={fillDirMenu} onClick={() => cmds.fill('down')} keytip="FI" />
          <StackSplit icon={<Delete20Regular />} label="Clear" tip={{ title: 'Clear' }} menu={clearMenu} onClick={() => cmds.clearSelection('contents')} keytip="E" />
        </div>
        <SplitButton icon={<ArrowSort20Regular />} label="Sort & Filter" tip={{ title: 'Sort & Filter' }} menu={sortMenu} menuOnly keytip="S" />
        <SplitButton icon={<Search20Regular />} label="Find & Select" tip={{ title: 'Find & Select' }} menu={findMenu} menuOnly keytip="FD" />
      </Group>


      {colorMenu && (
        <ColorMenu
          x={colorMenu.x || 300}
          y={colorMenu.y || 130}
          mode={colorMenu.mode}
          onClose={() => setColorMenu(null)}
          onPick={(c) => {
            if (colorMenu.mode === 'fill') { if (c) setLastFill(c); setStyle({ fill: c }, 'Fill Color') }
            else { if (c) setLastFontColor(c); setStyle({ color: c }, 'Font Color') }
          }}
        />
      )}
    </>
  )
}

function ColorSplit({
  icon, color, tip, onClick, onOpen, keytip
}: { icon: ReactNode; color: string; tip: { title: string }; onClick: () => void; onOpen: (x: number, y: number) => void; keytip?: string }) {
  const boxRef = useRef<HTMLDivElement | null>(null)
  return (
    <div className="xl-small-btn xl-split" ref={boxRef}>
      <button className="xl-small-split-main" title={tip.title} onClick={onClick}>
        <span className="xl-small-icon" style={{ position: 'relative' }}>{icon}<span className="xl-color-bar" style={{ background: color }} /><KeyTipBadge code={keytip} /></span>
      </button>
      <button
        className="xl-small-split-caret"
        onClick={() => {
          const r = boxRef.current!.getBoundingClientRect()
          onOpen(r.left, r.bottom)
        }}
      >
        <ChevronDown12Regular />
      </button>
    </div>
  )
}

/** Text-based icons for the Number group, drawn like Excel's ($, %, comma, decimal arrows). */
function Glyph({ text, sub }: { text: string; sub?: string }) {
  return (
    <span className="xl-glyph">
      {text}
      {sub && <small>{sub}</small>}
    </span>
  )
}

function DecimalGlyph({ more }: { more?: boolean }) {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1">
      <text x="1" y="8" fontSize="7" fill="currentColor" stroke="none">{more ? '.0' : '.00'}</text>
      <text x={more ? 8 : 10} y="18" fontSize="7" fill="currentColor" stroke="none">{more ? '.00' : '.0'}</text>
      <path d={more ? 'M2 14h5m-2-2 2 2-2 2' : 'M8 14H3m2-2-2 2 2 2'} stroke="#2b7cd3" />
    </svg>
  )
}

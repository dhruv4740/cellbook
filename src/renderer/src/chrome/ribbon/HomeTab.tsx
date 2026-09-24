// Home tab: Clipboard, Font, Alignment, Number, Styles, Cells, Editing, Add-ins.
import { ReactNode, useRef, useState } from 'react'
import { useStore } from '../../model/store'
import { useUi, setUi, getUi } from '../ui'
import { CellStyle, normRange } from '../../model/types'
import { MenuItem } from '../../components/Menu'
import { Group, GroupSep, BigButton, SplitButton, SmallButton, DialogLauncher, Combo } from './controls'
import { ChevronDown12Regular } from '@fluentui/react-icons'
import { ColorMenu } from './ColorMenu'
import * as cmds from '../../grid/commands'
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
  'Aptos Narrow', 'Aptos', 'Aptos Display', 'Arial', 'Arial Black', 'Bahnschrift', 'Calibri', 'Calibri Light',
  'Cambria', 'Candara', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Georgia',
  'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'
]
const SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72']

const NUMFMT: [string, string][] = [
  ['General', 'General'],
  ['Number', '0.00'],
  ['Currency', '$#,##0.00'],
  ['Accounting', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'],
  ['Short Date', 'm/d/yyyy'],
  ['Long Date', 'dddd, mmmm d, yyyy'],
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
  ['Currency', { numFmt: '$#,##0.00' }],
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

  const setStyle = (patch: Partial<CellStyle>, label = 'Format Cells') => store.applyStyle(sheet, ranges, patch, label)
  const toggle = (key: 'bold' | 'italic' | 'strike', label: string) => setStyle({ [key]: !style[key] } as Partial<CellStyle>, label)

  const applyBorders = (kind: string) => {
    ranges.forEach((rg) => {
      const n = normRange(rg)
      store.applyStyle(sheet, [rg], (s, r, c) => {
        const b: NonNullable<CellStyle['borders']> = { ...s.borders }
        const thin = { style: 'thin', color: '#000000' } as const
        const thick = { style: 'medium', color: '#000000' } as const
        const clearAll = () => { delete b.top; delete b.right; delete b.bottom; delete b.left }
        switch (kind) {
          case 'bottom': b.bottom = thin; break
          case 'top': b.top = thin; break
          case 'left': b.left = thin; break
          case 'right': b.right = thin; break
          case 'none': clearAll(); break
          case 'all': b.top = thin; b.bottom = thin; b.left = thin; b.right = thin; break
          case 'outside':
            if (r === n.r1) b.top = thin; if (r === n.r2) b.bottom = thin
            if (c === n.c1) b.left = thin; if (c === n.c2) b.right = thin
            break
          case 'thickOutside':
            if (r === n.r1) b.top = thick; if (r === n.r2) b.bottom = thick
            if (c === n.c1) b.left = thick; if (c === n.c2) b.right = thick
            break
          case 'bottomDouble': b.bottom = { style: 'double', color: '#000000' }; break
          case 'thickBottom': b.bottom = thick; break
          case 'topBottom': b.top = thin; b.bottom = thin; break
          case 'topThickBottom': b.top = thin; b.bottom = thick; break
        }
        return { ...s, borders: Object.keys(b).length ? b : undefined }
      }, 'Borders')
    })
  }

  const decimalAdjust = (delta: number) => {
    const cur = style.numFmt && style.numFmt !== 'General' ? style.numFmt : '0'
    const isPct = /%$/.test(cur)
    const isCur = /^\$/.test(cur)
    const m = /\.(0+)/.exec(cur)
    const dec = Math.max(0, (m ? m[1].length : 0) + delta)
    const base = isCur ? '$#,##0' : (/#,##0/.test(cur) ? '#,##0' : '0')
    setStyle({ numFmt: base + (dec > 0 ? '.' + '0'.repeat(dec) : '') + (isPct ? '%' : '') }, 'Number Format')
  }

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
    { label: 'No Border', onClick: () => applyBorders('none') },
    { label: '&All Borders', onClick: () => applyBorders('all') },
    { label: 'Ou&tside Borders', onClick: () => applyBorders('outside') },
    { label: '&Thick Outside Borders', onClick: () => applyBorders('thickOutside') },
    { label: 'Botto&m Double Border', onClick: () => applyBorders('bottomDouble') },
    { label: 'Thick Botto&m Border', onClick: () => applyBorders('thickBottom') },
    { label: 'T&op and Bottom Border', onClick: () => applyBorders('topBottom') },
    { label: 'Top and Thick Bottom Border', onClick: () => applyBorders('topThickBottom') },
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
  const formatMenu: MenuItem[] = [
    { label: 'Row &Height...', onClick: () => setUi({ dialog: { kind: 'rowHeight' } }) },
    { label: '&AutoFit Row Height', onClick: () => cmds.autoFitRows() },
    { label: 'Column &Width...', onClick: () => setUi({ dialog: { kind: 'columnWidth' } }) },
    { label: 'AutoFit Column &Width', onClick: () => cmds.autoFitColumns() },
    { label: 'Default Width...', onClick: () => setUi({ dialog: { kind: 'columnWidth' } }) },
    'separator',
    { label: 'Hide &Rows', onClick: () => cmds.hideRows(true) },
    { label: 'Unhide Rows', onClick: () => cmds.hideRows(false) },
    { label: 'Hide &Columns', onClick: () => cmds.hideColumns(true) },
    { label: 'Unhide Columns', onClick: () => cmds.hideColumns(false) },
    { label: 'Hide &Sheet', onClick: () => {} },
    'separator',
    { label: '&Rename Sheet', onClick: () => setUi({ dialog: { kind: 'renameSheet' } }) },
    { label: '&Move or Copy Sheet...', onClick: () => {} },
    { label: 'Tab &Color', onClick: () => {} },
    'separator',
    { label: '&Protect Sheet...', onClick: () => {} },
    { label: 'Lock Cell', onClick: () => setStyle({ locked: !style.locked }, 'Lock Cell') },
    { label: 'Format Cells...', shortcut: 'Ctrl+1', onClick: () => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } }) }
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
    { label: '&Series...', onClick: () => {} },
    { label: '&Justify', onClick: () => {} }
  ]
  const clearMenu: MenuItem[] = [
    { label: 'Clear &All', onClick: () => cmds.clearSelection('all') },
    { label: 'Clear &Formats', onClick: () => cmds.clearSelection('formats') },
    { label: 'Clear &Contents', shortcut: 'Del', onClick: () => cmds.clearSelection('contents') },
    { label: 'Clear Co&mments and Notes', onClick: () => {} },
    { label: 'Clear H&yperlinks', onClick: () => {} }
  ]
  const sortMenu: MenuItem[] = [
    { label: 'Sort &A to Z', onClick: () => cmds.sortSelection('asc') },
    { label: 'Sort &Z to A', onClick: () => cmds.sortSelection('desc') },
    { label: 'Custom Sort...', onClick: () => {} },
    'separator',
    { label: '&Filter', icon: <Filter20Regular />, onClick: () => {} },
    { label: 'Clear', onClick: () => {} },
    { label: 'Reapply', onClick: () => {} }
  ]
  const findMenu: MenuItem[] = [
    { label: '&Find...', shortcut: 'Ctrl+F', onClick: () => {} },
    { label: '&Replace...', shortcut: 'Ctrl+H', onClick: () => {} },
    { label: 'Go &To...', shortcut: 'Ctrl+G', onClick: () => {} },
    { label: 'Go To &Special...', onClick: () => {} },
    'separator',
    { label: 'Select Objects', onClick: () => {} }
  ]

  return (
    <>
      <Group label="Clipboard">
        <div className="xl-col">
          <SplitButton icon={<ClipboardPaste24Regular fontSize={28} />} label="Paste" tip={{ title: 'Paste (Ctrl+V)', desc: 'Paste the contents of the Clipboard.' }} onClick={() => cmds.paste('all')} menu={pasteMenu} />
        </div>
        <div className="xl-col xl-stack">
          <SmallButton icon={<Cut20Regular />} label="Cut" tip={{ title: 'Cut (Ctrl+X)' }} onClick={() => cmds.cut()} />
          <SmallButton icon={<Copy20Regular />} label="Copy" menu={[{ label: '&Copy', onClick: () => cmds.copy() }, { label: 'Copy as &Picture...', onClick: () => {} }]} tip={{ title: 'Copy (Ctrl+C)' }} onClick={() => cmds.copy()} />
          <SmallButton
            icon={<PaintBrush20Regular />}
            label="Format Painter"
            tip={{ title: 'Format Painter', desc: 'Copy formatting to another selection. Double-click to keep it on.' }}
            active={!!getUi().formatPainter}
            onClick={() => setUi({ formatPainter: getUi().formatPainter ? null : { once: true } })}
            onDoubleClick={() => setUi({ formatPainter: { once: false } })}
          />
        </div>
      </Group>
      <GroupSep />

      <Group label="Font">
        <div className="xl-col">
          <div className="xl-row">
            <Combo value={style.fontFamily ?? 'Aptos Narrow'} options={FONTS} width={120} tip={{ title: 'Font' }} onChange={(v) => setStyle({ fontFamily: v }, 'Font')} />
            <Combo value={String(style.fontSize ?? 11)} options={SIZES} width={48} tip={{ title: 'Font Size' }} onChange={(v) => { const n = +v; if (n > 0) setStyle({ fontSize: n }, 'Font Size') }} />
            <SmallButton icon={<FontIncrease20Regular />} tip={{ title: 'Increase Font Size' }} onClick={() => setStyle({ fontSize: (style.fontSize ?? 11) + 1 }, 'Font Size')} />
            <SmallButton icon={<FontDecrease20Regular />} tip={{ title: 'Decrease Font Size' }} onClick={() => setStyle({ fontSize: Math.max(1, (style.fontSize ?? 11) - 1) }, 'Font Size')} />
          </div>
          <div className="xl-row">
            <SmallButton icon={<TextBold20Regular />} active={!!style.bold} tip={{ title: 'Bold (Ctrl+B)' }} onClick={() => toggle('bold', 'Bold')} />
            <SmallButton icon={<TextItalic20Regular />} active={!!style.italic} tip={{ title: 'Italic (Ctrl+I)' }} onClick={() => toggle('italic', 'Italic')} />
            <SmallButton icon={<TextUnderline20Regular />} active={!!style.underline} tip={{ title: 'Underline (Ctrl+U)' }} menu={underlineMenu} onClick={() => setStyle({ underline: style.underline ? undefined : 'single' }, 'Underline')} />
            <SmallButton icon={<TextStrikethrough20Regular />} active={!!style.strike} tip={{ title: 'Strikethrough' }} onClick={() => toggle('strike', 'Strikethrough')} />
            <SmallButton icon={<BorderAll20Regular />} tip={{ title: 'Borders' }} menu={bordersMenu} onClick={() => applyBorders('bottom')} />
            <ColorSplit
              icon={<PaintBucket20Regular />}
              color={lastFill}
              tip={{ title: 'Fill Color' }}
              onClick={() => setStyle({ fill: lastFill }, 'Fill Color')}
              onOpen={(x, y) => setColorMenu({ x, y, mode: 'fill' })}
            />
            <ColorSplit
              icon={<TextFont20Regular />}
              color={style.color ?? lastFontColor}
              tip={{ title: 'Font Color' }}
              onClick={() => setStyle({ color: lastFontColor }, 'Font Color')}
              onOpen={(x, y) => setColorMenu({ x, y, mode: 'font' })}
            />
          </div>
        </div>
        <DialogLauncher tip={{ title: 'Format Cells: Font', shortcut: 'Ctrl+Shift+F' }} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Font' } } })} />
      </Group>
      <GroupSep />

      <Group label="Alignment">
        <div className="xl-col">
          <div className="xl-row">
            <SmallButton icon={<AlignTop20Regular />} active={style.vAlign === 'top'} tip={{ title: 'Top Align' }} onClick={() => setStyle({ vAlign: 'top' }, 'Align')} />
            <SmallButton icon={<AlignCenterVertical20Regular />} active={style.vAlign === 'middle'} tip={{ title: 'Middle Align' }} onClick={() => setStyle({ vAlign: 'middle' }, 'Align')} />
            <SmallButton icon={<AlignBottom20Regular />} active={style.vAlign === 'bottom'} tip={{ title: 'Bottom Align' }} onClick={() => setStyle({ vAlign: 'bottom' }, 'Align')} />
            <SmallButton icon={<TextAlignCenterRotate27020Regular />} tip={{ title: 'Orientation' }} menu={orientationMenu} onClick={() => {}} />
          </div>
          <div className="xl-row">
            <SmallButton icon={<AlignLeft20Regular />} active={style.hAlign === 'left'} tip={{ title: 'Align Left' }} onClick={() => setStyle({ hAlign: 'left' }, 'Align')} />
            <SmallButton icon={<AlignCenterHorizontal20Regular />} active={style.hAlign === 'center'} tip={{ title: 'Center' }} onClick={() => setStyle({ hAlign: 'center' }, 'Align')} />
            <SmallButton icon={<AlignRight20Regular />} active={style.hAlign === 'right'} tip={{ title: 'Align Right' }} onClick={() => setStyle({ hAlign: 'right' }, 'Align')} />
            <SmallButton icon={<TextWrap20Regular />} active={!!style.wrap} tip={{ title: 'Wrap Text' }} onClick={() => setStyle({ wrap: !style.wrap }, 'Wrap Text')} />
          </div>
          <div className="xl-row">
            <SmallButton icon={<TextIndentDecrease20Regular />} tip={{ title: 'Decrease Indent' }} onClick={() => setStyle({ indent: Math.max(0, (style.indent ?? 0) - 1) }, 'Indent')} />
            <SmallButton icon={<TextIndentIncrease20Regular />} tip={{ title: 'Increase Indent' }} onClick={() => setStyle({ indent: (style.indent ?? 0) + 1 }, 'Indent')} />
            <SplitButton icon={<TableCellsMerge20Regular />} label="Merge & Center" tip={{ title: 'Merge & Center' }} onClick={() => cmds.mergeCells('center')} menu={mergeMenu} />
          </div>
        </div>
        <DialogLauncher tip={{ title: 'Format Cells: Alignment' }} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Alignment' } } })} />
      </Group>
      <GroupSep />

      <Group label="Number" style={{ width: 130 }}>
        <div className="xl-col">
          <Combo value={curFmtName} options={NUMFMT.map(([n]) => n)} width={122} editable={false} tip={{ title: 'Number Format' }} onPick={(name) => { const f = NUMFMT.find(([n]) => n === name); if (f) setStyle({ numFmt: f[1] }, 'Number Format') }} />
          <div className="xl-row">
            <SmallButton icon={<TextPercent20Regular />} tip={{ title: 'Accounting Number Format' }} menu={[{ label: 'Accounting', onClick: () => setStyle({ numFmt: NUMFMT[3][1] }, 'Number Format') }, { label: '$ English (US)', onClick: () => setStyle({ numFmt: '$#,##0.00' }, 'Number Format') }]} onClick={() => setStyle({ numFmt: NUMFMT[3][1] }, 'Number Format')} />
            <SmallButton icon={<TextPercent20Regular />} tip={{ title: 'Percent Style (Ctrl+Shift+%)' }} onClick={() => setStyle({ numFmt: '0%' }, 'Number Format')} />
            <SmallButton icon={<ArrowSort20Regular />} tip={{ title: 'Comma Style' }} onClick={() => setStyle({ numFmt: '#,##0.00' }, 'Number Format')} />
            <SmallButton icon={<TextIndentIncrease20Regular />} tip={{ title: 'Increase Decimal' }} onClick={() => decimalAdjust(1)} />
            <SmallButton icon={<TextIndentDecrease20Regular />} tip={{ title: 'Decrease Decimal' }} onClick={() => decimalAdjust(-1)} />
          </div>
        </div>
        <DialogLauncher tip={{ title: 'Format Cells: Number' }} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } })} />
      </Group>
      <GroupSep />

      <Group label="Styles">
        <BigButton icon={<TableCellsMerge20Regular fontSize={28} />} label="Conditional Formatting" tip={{ title: 'Conditional Formatting' }} menu={[{ label: 'Highlight Cells Rules', submenu: [{ label: 'Greater Than...', onClick: () => {} }] }, { label: 'Clear Rules', onClick: () => {} }, { label: 'Manage Rules...', onClick: () => {} }]} onClick={() => {}} />
        <BigButton icon={<TableSimple20Regular fontSize={28} />} label="Format as Table" tip={{ title: 'Format as Table' }} menu={[{ label: 'New Table Style...', onClick: () => {} }]} onClick={() => {}} />
        <BigButton
          icon={<TableInsertRow20Regular fontSize={28} />}
          label="Cell Styles"
          tip={{ title: 'Cell Styles' }}
          onClick={() => {}}
          menu={CELL_STYLES.map(([name, s]) => ({ label: name, onClick: () => setStyle(s, 'Cell Styles') } as MenuItem))}
        />
      </Group>
      <GroupSep />

      <Group label="Cells">
        <div className="xl-col xl-stack">
          <SplitButton icon={<TableInsertRow20Regular />} label="Insert" tip={{ title: 'Insert Cells' }} menu={insertMenu} onClick={() => cmds.insertCellsDialog()} />
          <SplitButton icon={<Delete20Regular />} label="Delete" tip={{ title: 'Delete Cells' }} menu={deleteMenu} onClick={() => cmds.deleteCellsDialog()} />
          <SplitButton icon={<ArrowExpand20Regular />} label="Format" tip={{ title: 'Format' }} menu={formatMenu} onClick={() => setUi({ dialog: { kind: 'formatCells', props: { tab: 'Number' } } })} />
        </div>
      </Group>
      <GroupSep />

      <Group label="Editing">
        <div className="xl-col xl-stack">
          <SplitButton icon={<MathFormula20Regular />} label="AutoSum" tip={{ title: 'AutoSum (Alt+=)' }} menu={autoSumMenu} onClick={() => cmds.autoSum('SUM')} />
          <SplitButton icon={<TableInsertRow20Regular />} label="Fill" tip={{ title: 'Fill' }} menu={fillDirMenu} onClick={() => cmds.fill('down')} />
          <SplitButton icon={<Delete20Regular />} label="Clear" tip={{ title: 'Clear' }} menu={clearMenu} onClick={() => cmds.clearSelection('contents')} />
        </div>
        <div className="xl-col xl-stack">
          <SplitButton icon={<ArrowSort20Regular />} label="Sort & Filter" tip={{ title: 'Sort & Filter' }} menu={sortMenu} onClick={() => cmds.sortSelection('asc')} />
          <SplitButton icon={<Search20Regular />} label="Find & Select" tip={{ title: 'Find & Select' }} menu={findMenu} onClick={() => {}} />
        </div>
      </Group>
      <GroupSep />

      <Group label="Add-ins">
        <BigButton icon={<DataFunnel20Regular fontSize={28} />} label="Get Add-ins" tip={{ title: 'Get Add-ins' }} onClick={() => {}} />
      </Group>
      <GroupSep />
      <Group label="Analysis">
        <BigButton icon={<DataHistogram20Regular fontSize={28} />} label="Analyze Data" tip={{ title: 'Analyze Data' }} onClick={() => {}} />
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
  icon, color, tip, onClick, onOpen
}: { icon: ReactNode; color: string; tip: { title: string }; onClick: () => void; onOpen: (x: number, y: number) => void }) {
  const boxRef = useRef<HTMLDivElement | null>(null)
  return (
    <div className="xl-small-btn xl-split" ref={boxRef}>
      <button className="xl-small-split-main" title={tip.title} onClick={onClick}>
        <span className="xl-small-icon">{icon}<span className="xl-color-bar" style={{ background: color }} /></span>
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

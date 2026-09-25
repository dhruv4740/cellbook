// Table Design contextual tab (KeyTip JT): shown whenever the active cell is inside a Table.
import { useEffect, useRef, useState } from 'react'
import { useStore, store } from '../model/store'
import { setUi } from '../chrome/ui'
import { Group, GroupSep, SmallButton, KeyTipBadge } from '../chrome/ribbon/controls'
import { ChevronDown12Regular } from '@fluentui/react-icons'
import { TableStyleGallery } from './TableStyleGallery'
import { tableAt, setTableOption, renameTable } from '../model/tables'
import { TableStyleOptions } from '../model/types'
import './TableDesignTab.css'
import {
  TableSimple20Regular, ArrowExpand20Regular, DataHistogram20Regular, TableDeleteRow20Regular,
  Filter20Regular, ArrowSync20Regular, Share20Regular, TableSettings20Regular
} from '@fluentui/react-icons'

export default function TableDesignTab() {
  const s = useStore()
  const sheet = s.active
  const active = sheet.selection.active
  const table = tableAt(sheet, active.r, active.c)
  const [gallery, setGallery] = useState<{ x: number; y: number } | null>(null)
  const [nameText, setNameText] = useState(table?.name ?? '')
  const [nameError, setNameError] = useState<string | null>(null)

  useEffect(() => { if (table) setNameText(table.name) }, [table?.name])

  if (!table) return null // deselected mid-render (sheet switch); Ribbon falls back to Home next tick

  const commitName = () => {
    if (nameText === table.name) return
    const err = renameTable(store, sheet, table, nameText)
    if (err) { setNameError(err); setNameText(table.name) } else setNameError(null)
  }

  const opt = (key: keyof TableStyleOptions, label: string, keytip?: string) => (
    <SmallButton
      label={label}
      icon={<span className="xl-checkbox" aria-hidden>{table.options[key] ? '☑' : '☐'}</span>}
      active={table.options[key]}
      tip={{ title: label }}
      onClick={() => setTableOption(store, sheet, table, key, !table.options[key])}
      keytip={keytip}
    />
  )

  return (
    <>
      <Group label="Properties">
        <div className="xl-col">
          <div style={{ fontSize: 11 }}>Table Name:</div>
          <input
            className="xl-tdt-name"
            value={nameText}
            title={nameError ?? undefined}
            onChange={(e) => setNameText(e.target.value)}
            onBlur={commitName}
            onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
          />
        </div>
        <div className="xl-col">
          <SmallButton icon={<ArrowExpand20Regular />} label="Resize Table" tip={{ title: 'Resize Table' }} onClick={() => setUi({ dialog: { kind: 'resizeTable', props: { table } } })} keytip="RT" />
          <SmallButton icon={<DataHistogram20Regular />} label="Summarize with PivotTable" disabled tip={{ title: 'Summarize with PivotTable (not implemented)' }} onClick={() => {}} />
        </div>
      </Group>
      <GroupSep />
      <Group label="Tools">
        <div className="xl-col xl-stack">
          <SmallButton icon={<TableDeleteRow20Regular />} label="Remove Duplicates" tip={{ title: 'Remove Duplicates' }} onClick={() => { store.setSelection({ ranges: [table.range], active: { r: table.range.r1, c: table.range.c1 } }, sheet); setUi({ dialog: { kind: 'removeDuplicates' } }) }} keytip="M" />
          <SmallButton icon={<Filter20Regular />} label="Insert Slicer" disabled tip={{ title: 'Insert Slicer (not implemented)' }} onClick={() => {}} />
          <SmallButton icon={<TableSimple20Regular />} label="Convert to Range" tip={{ title: 'Convert to Range' }} onClick={() => setUi({ dialog: { kind: 'convertTableToRange', props: { table } } })} keytip="CV" />
        </div>
      </Group>
      <GroupSep />
      <Group label="External Table Data">
        <div className="xl-col xl-stack">
          <SmallButton icon={<Share20Regular />} label="Export" disabled tip={{ title: 'Export (no external data source)' }} onClick={() => {}} />
          <SmallButton icon={<ArrowSync20Regular />} label="Refresh" disabled tip={{ title: 'Refresh (no external data source)' }} onClick={() => {}} />
          <SmallButton icon={<TableSettings20Regular />} label="Unlink" disabled tip={{ title: 'Unlink (no external data source)' }} onClick={() => {}} />
        </div>
      </Group>
      <GroupSep />
      <Group label="Table Style Options">
        <div className="xl-col">
          <div className="xl-row">{opt('headerRow', 'Header Row', 'HR')}{opt('firstColumn', 'First Column', 'FC')}{opt('filterButton', 'Filter Button', 'FB')}</div>
          <div className="xl-row">{opt('totalRow', 'Total Row', 'TR')}{opt('lastColumn', 'Last Column', 'LC')}</div>
          <div className="xl-row">{opt('bandedRows', 'Banded Rows', 'BR')}{opt('bandedColumns', 'Banded Columns', 'BC')}</div>
        </div>
      </Group>
      <GroupSep />
      <Group label="Table Styles">
        <TableStylesButton table={table} onOpen={(x, y) => setGallery({ x, y })} />
      </Group>

      {gallery && (
        <TableStyleGallery
          x={gallery.x}
          y={gallery.y}
          current={table.styleName}
          onClose={() => setGallery(null)}
          onPick={(styleName) => { setGallery(null); import('../model/tables').then((m) => m.setTableStyle(store, sheet, table, styleName)) }}
        />
      )}
    </>
  )
}

function TableStylesButton({ table, onOpen }: { table: { styleName: string }; onOpen: (x: number, y: number) => void }) {
  const ref = useRef<HTMLButtonElement>(null)
  return (
    <button
      ref={ref}
      className="xl-big-btn wide"
      onClick={() => { const r = ref.current!.getBoundingClientRect(); onOpen(r.left, r.bottom) }}
    >
      <span className="xl-big-icon" style={{ position: 'relative' }}>
        <TableSimple20Regular fontSize={28} />
        <KeyTipBadge code="TS" />
      </span>
      <span className="xl-big-label">{table.styleName.replace('TableStyle', '')}<ChevronDown12Regular className="xl-drop-caret" /></span>
    </button>
  )
}

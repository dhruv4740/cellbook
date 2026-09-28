// Sheet navigation strip: prev/next arrows, tabs, new-sheet button, drag reorder, right-click menu.
import { useEffect, useRef, useState } from 'react'
import { hideSheet } from './sheetCommands'
import { useStore } from '../model/store'
import { useUi, setUi, notAvailable } from './ui'
import { Menu, MenuItem, useMenu } from '../components/Menu'
import { Dialog } from '../components/Dialog'
import { ChevronLeft20Regular, ChevronRight20Regular, Add20Regular } from '@fluentui/react-icons'
import './SheetTabs.css'
import { toggleSheetProtection } from '../dialogs/ProtectDialogs'

export default function SheetTabs() {
  const store = useStore()
  const ui = useUi()
  const grouped = ui.groupedSheets && ui.groupedSheets.length > 1 ? ui.groupedSheets : null
  const [menu, openMenu, closeMenu] = useMenu()
  const [menuIndex, setMenuIndex] = useState(0)
  const [renaming, setRenaming] = useState<number | null>(null)
  const [renameValue, setRenameValue] = useState('')
  const [renameError, setRenameError] = useState<string | null>(null)
  const [colorMenuAt, setColorMenuAt] = useState<{ x: number; y: number } | null>(null)
  const dragFrom = useRef<number | null>(null)
  const [dragOver, setDragOver] = useState<number | null>(null)

  const startRename = (i: number) => {
    setRenaming(i)
    setRenameValue(store.sheets[i].name)
  }
  const commitRename = () => {
    if (renaming === null) return
    const err = store.renameSheet(renaming, renameValue)
    if (err) { setRenameError(err); return }
    setRenaming(null)
  }

  // Home > Format > Rename Sheet puts the active tab into rename mode, like Excel.
  useEffect(() => {
    if (ui.dialog?.kind === 'renameSheet') { setUi({ dialog: null }); startRename(store.activeIndex) }
  })

  const items = (i: number): MenuItem[] => [
    { label: '&Insert...', disabled: !!store.bookProtection?.lockStructure, onClick: () => store.addSheet(undefined, i) },
    { label: '&Delete', disabled: !!store.bookProtection?.lockStructure, onClick: () => store.deleteSheet(i) },
    { label: '&Rename', disabled: !!store.bookProtection?.lockStructure, onClick: () => startRename(i) },
    { label: '&Move or Copy...', disabled: !!store.bookProtection?.lockStructure, onClick: () => setUi({ dialog: { kind: 'moveCopySheet', props: { index: i } } }) },
    'separator',
    { label: 'View &Code', onClick: () => notAvailable('View Code') },
    { label: store.sheets[i].layout.protection ? 'Unprotect S&heet...' : '&Protect Sheet...', onClick: () => toggleSheetProtection(store.sheets[i]) },
    {
      label: 'Ta&b Color',
      submenu: ['#FF0000', '#FFC000', '#FFFF00', '#00B050', '#00B0F0', '#0070C0', '#7030A0', 'No Color'].map((c) => ({
        label: c === 'No Color' ? 'No Color' : c,
        icon: c === 'No Color' ? undefined : <span style={{ display: 'inline-block', width: 12, height: 12, background: c }} />,
        onClick: () => store.updateLayout(store.sheets[i], 'Tab Color', (l) => (l.tabColor = c === 'No Color' ? undefined : c))
      }))
    },
    { label: '&Hide', onClick: () => hideSheet(i) },
    { label: 'Unhide...', disabled: !store.sheets.some((s) => s.layout.sheetHidden === 'hidden'), onClick: () => setUi({ dialog: { kind: 'unhideSheet' } }) },
    'separator',
    { label: 'Select All Sheets', onClick: () => setUi({ groupedSheets: store.sheets.map((_, k) => k).filter((k) => !store.sheets[k].layout.sheetHidden) }) }
  ]

  return (
    <div className="xl-sheettabs">
      <div className="xl-sheet-nav">
        <button className="xl-sheet-navbtn" title="Activate First Sheet" onClick={() => store.setActiveSheet(store.sheets.findIndex((s) => !s.layout.sheetHidden))}><ChevronLeft20Regular /></button>
        <button className="xl-sheet-navbtn" title="Activate Previous Sheet" onClick={() => { for (let k = store.activeIndex - 1; k >= 0; k--) if (!store.sheets[k].layout.sheetHidden) { store.setActiveSheet(k); break } }}><ChevronRight20Regular style={{ transform: 'rotate(180deg)' }} /></button>
        <button className="xl-sheet-navbtn" title="Activate Next Sheet" onClick={() => { for (let k = store.activeIndex + 1; k < store.sheets.length; k++) if (!store.sheets[k].layout.sheetHidden) { store.setActiveSheet(k); break } }}><ChevronRight20Regular /></button>
      </div>
      <div className="xl-sheet-list">
        {store.sheets.map((s, i) => s.layout.sheetHidden ? null : (
          <div
            key={s.uid}
            className={'xl-sheet-tab' + (i === store.activeIndex ? ' active' : '') + (dragOver === i ? ' drag-over' : '') + (grouped?.includes(i) ? ' grouped' : '')}
            draggable
            onDragStart={() => (dragFrom.current = i)}
            onDragOver={(e) => { e.preventDefault(); setDragOver(i) }}
            onDragLeave={() => setDragOver((d) => (d === i ? null : d))}
            onDrop={(e) => { e.preventDefault(); if (dragFrom.current !== null && dragFrom.current !== i) store.moveSheet(dragFrom.current, i); dragFrom.current = null; setDragOver(null) }}
            onClick={(e) => { if (!e.shiftKey && !e.ctrlKey && !e.metaKey) setUi({ groupedSheets: null }); store.setActiveSheet(i) }}
            onDoubleClick={() => startRename(i)}
            onContextMenu={(e) => { e.preventDefault(); setMenuIndex(i); openMenu(e) }}
            style={s.layout.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.layout.tabColor}` } : undefined}
          >
            {renaming === i ? (
              <input
                autoFocus
                className="xl-sheet-rename"
                value={renameValue}
                onChange={(e) => setRenameValue(e.target.value)}
                onBlur={commitRename}
                onKeyDown={(e) => { if (e.key === 'Enter') commitRename(); if (e.key === 'Escape') setRenaming(null) }}
                onFocus={(e) => e.currentTarget.select()}
              />
            ) : (
              s.name
            )}
          </div>
        ))}
        <button className="xl-sheet-add" title="New Sheet (Shift+F11)" onClick={() => store.addSheet()}><Add20Regular /></button>
      </div>
      <div className="xl-sheet-scrollarea" />

      {menu && <Menu x={menu.x} y={menu.y} items={items(menuIndex)} onClose={closeMenu} />}
      {renameError && (
        <Dialog title="Microsoft Excel" onCancel={() => setRenameError(null)} cancelLabel="OK">
          <div style={{ whiteSpace: 'pre-wrap', fontSize: 12, maxWidth: 320 }}>{renameError}</div>
        </Dialog>
      )}
    </div>
  )
}

// Sheet navigation strip: prev/next arrows, tabs, new-sheet button, drag reorder, right-click menu.
import { useRef, useState } from 'react'
import { useStore } from '../model/store'
import { Menu, MenuItem, useMenu } from '../components/Menu'
import { Dialog } from '../components/Dialog'
import { ChevronLeft20Regular, ChevronRight20Regular, Add20Regular } from '@fluentui/react-icons'
import './SheetTabs.css'

export default function SheetTabs() {
  const store = useStore()
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

  const items = (i: number): MenuItem[] => [
    { label: '&Insert...', onClick: () => store.addSheet(undefined, i) },
    { label: '&Delete', onClick: () => store.deleteSheet(i) },
    { label: '&Rename', onClick: () => startRename(i) },
    { label: '&Move or Copy...', onClick: () => {} },
    'separator',
    { label: 'View &Code', onClick: () => {} },
    { label: '&Protect Sheet...', onClick: () => {} },
    {
      label: 'Ta&b Color',
      submenu: ['#FF0000', '#FFC000', '#FFFF00', '#00B050', '#00B0F0', '#0070C0', '#7030A0', 'No Color'].map((c) => ({
        label: c === 'No Color' ? 'No Color' : c,
        icon: c === 'No Color' ? undefined : <span style={{ display: 'inline-block', width: 12, height: 12, background: c }} />,
        onClick: () => store.updateLayout(store.sheets[i], 'Tab Color', (l) => (l.tabColor = c === 'No Color' ? undefined : c))
      }))
    },
    { label: '&Hide', onClick: () => {} },
    { label: 'Unhide...', onClick: () => {} },
    'separator',
    { label: 'Select All Sheets', onClick: () => {} }
  ]

  return (
    <div className="xl-sheettabs">
      <div className="xl-sheet-nav">
        <button className="xl-sheet-navbtn" title="Activate First Sheet" onClick={() => store.setActiveSheet(0)}><ChevronLeft20Regular /></button>
        <button className="xl-sheet-navbtn" title="Activate Previous Sheet" onClick={() => store.setActiveSheet(Math.max(0, store.activeIndex - 1))}><ChevronRight20Regular style={{ transform: 'rotate(180deg)' }} /></button>
        <button className="xl-sheet-navbtn" title="Activate Next Sheet" onClick={() => store.setActiveSheet(Math.min(store.sheets.length - 1, store.activeIndex + 1))}><ChevronRight20Regular /></button>
      </div>
      <div className="xl-sheet-list">
        {store.sheets.map((s, i) => (
          <div
            key={s.uid}
            className={'xl-sheet-tab' + (i === store.activeIndex ? ' active' : '') + (dragOver === i ? ' drag-over' : '')}
            draggable
            onDragStart={() => (dragFrom.current = i)}
            onDragOver={(e) => { e.preventDefault(); setDragOver(i) }}
            onDragLeave={() => setDragOver((d) => (d === i ? null : d))}
            onDrop={(e) => { e.preventDefault(); if (dragFrom.current !== null && dragFrom.current !== i) store.moveSheet(dragFrom.current, i); dragFrom.current = null; setDragOver(null) }}
            onClick={() => store.setActiveSheet(i)}
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

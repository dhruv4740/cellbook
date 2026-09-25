// Frameless window title bar: logo, AutoSave, Quick Access Toolbar, file name, search, avatar, window controls.
import { save, isAutoSaveOn, onAutoSaveChange, setAutoSave } from '../io/fileOps'
import { useEffect, useRef, useState } from 'react'
import { useStore } from '../model/store'
import { useUi } from './ui'
import { Menu, MenuItem } from '../components/Menu'
import {
  Save20Regular, ArrowUndo20Regular, ArrowRedo20Regular, ChevronDown12Regular,
  Search20Regular, PersonCircle24Regular, Dismiss20Regular, Square20Regular, ArrowMinimize20Regular,
  SquareMultiple16Regular
} from '@fluentui/react-icons'
import './TitleBar.css'

export default function TitleBar() {
  const store = useStore()
  const ui = useUi()
  const grouped = !!(ui.groupedSheets && ui.groupedSheets.length > 1)
  const [autoSave, setAutoSaveState] = useState(isAutoSaveOn())
  useEffect(() => onAutoSaveChange(() => setAutoSaveState(isAutoSaveOn())), [])
  const [maximized, setMaximized] = useState(false)
  const [undoMenu, setUndoMenu] = useState<{ x: number; y: number } | null>(null)
  const [qatMenu, setQatMenu] = useState<{ x: number; y: number } | null>(null)
  const [search, setSearch] = useState('')
  const undoBtn = useRef<HTMLDivElement>(null)
  const qatBtn = useRef<HTMLButtonElement>(null)
  const [viewIndex, setViewIndex] = useState(0)

  useEffect(() => {
    return window.api?.onViewIndexChanged?.((idx) => setViewIndex(idx))
  }, [])

  const viewSuffix = viewIndex > 0 ? `:${viewIndex}` : ''
  const title = store.fileName + viewSuffix + (grouped ? ' [Group]' : '') + ' - Excel'
  useEffect(() => {
    document.title = title
    window.api?.setTitle?.(title)
  }, [title])
  useEffect(() => {
    window.api?.onMaximized?.((m) => setMaximized(m))
  }, [])

  const undoItems: MenuItem[] = store.undoLabels.length
    ? store.undoLabels.map((l, i) => ({ label: l, onClick: () => { for (let k = 0; k <= i; k++) store.undo() } }))
    : [{ label: 'Nothing to undo', disabled: true }]

  return (
    <div className="xl-titlebar">
      <div className="xl-titlebar-drag">
        <div className="xl-logo">
          <svg width="20" height="20" viewBox="0 0 20 20"><rect width="20" height="20" rx="3" fill="#185C37" /><path d="M4 4h6.5l5.5 4-5.5 4L4 16V4z" fill="none" /><text x="3.2" y="15.3" fontFamily="Segoe UI" fontWeight="700" fontSize="12" fill="#fff">X</text></svg>
        </div>
        <button className={'xl-autosave' + (autoSave ? ' on' : '')} onClick={() => void setAutoSave(!autoSave)} title={autoSave ? 'AutoSave is on: changes are saved automatically' : 'Turn on AutoSave to save changes automatically'}>
          <span className="xl-autosave-toggle" />
          AutoSave <b>{autoSave ? 'On' : 'Off'}</b>
        </button>

        <div className="xl-qat">
          <button className="xl-qat-btn" title="Save (Ctrl+S)" onClick={() => void save()} style={{ position: 'relative' }}><Save20Regular /><span className="xl-keytip-badge" data-keytip="1">1</span></button>
          <div className="xl-qat-split" ref={undoBtn}>
            <button className="xl-qat-btn" title="Undo (Ctrl+Z)" disabled={!store.canUndo} onClick={() => store.undo()} style={{ position: 'relative' }}><ArrowUndo20Regular /><span className="xl-keytip-badge" data-keytip="2">2</span></button>
            <button
              className="xl-qat-caret"
              disabled={!store.canUndo}
              onClick={() => { const r = undoBtn.current!.getBoundingClientRect(); setUndoMenu({ x: r.left, y: r.bottom }) }}
            >
              <ChevronDown12Regular />
            </button>
          </div>
          <button className="xl-qat-btn" title="Redo (Ctrl+Y)" disabled={!store.canRedo} onClick={() => store.redo()} style={{ position: 'relative' }}><ArrowRedo20Regular /><span className="xl-keytip-badge" data-keytip="3">3</span></button>
          <button ref={qatBtn} className="xl-qat-btn" title="Customize Quick Access Toolbar" onClick={() => { const r = qatBtn.current!.getBoundingClientRect(); setQatMenu({ x: r.left, y: r.bottom }) }}>
            <ChevronDown12Regular />
          </button>
        </div>

        <div className="xl-titlebar-center">
          <span className="xl-filename">{store.fileName}{viewSuffix}{grouped ? ' [Group]' : ''}</span>
          <span className="xl-saved">{store.dirty ? '' : '• Saved to this PC'}</span>
        </div>

        <div className="xl-titlebar-search">
          <Search20Regular />
          <input placeholder="Search (Alt+Q)" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
      </div>

      <div className="xl-titlebar-right">
        <button className="xl-avatar" title="dhruvkothari13062006@gmail.com"><PersonCircle24Regular /></button>
        <div className="xl-winctl">
          <button className="xl-winbtn" onClick={() => window.api?.windowControl?.('minimize')}><ArrowMinimize20Regular /></button>
          <button className="xl-winbtn" onClick={() => window.api?.windowControl?.('maximize')}>
            {maximized ? <SquareMultiple16Regular /> : <Square20Regular style={{ transform: 'scale(0.7)' }} />}
          </button>
          <button className="xl-winbtn close" onClick={() => window.api?.windowControl?.('close')}><Dismiss20Regular /></button>
        </div>
      </div>

      {undoMenu && <Menu x={undoMenu.x} y={undoMenu.y} items={undoItems} onClose={() => setUndoMenu(null)} />}
      {qatMenu && (
        <Menu
          x={qatMenu.x}
          y={qatMenu.y}
          onClose={() => setQatMenu(null)}
          items={[
            { label: 'Customize Quick Access Toolbar', disabled: true },
            'separator',
            { label: 'New', checked: false, onClick: () => {} },
            { label: 'Open', checked: false, onClick: () => {} },
            { label: 'Save', checked: true, onClick: () => {} },
            { label: 'Email', checked: false, onClick: () => {} },
            { label: 'Quick Print', checked: false, onClick: () => {} },
            { label: 'Print Preview and Print', checked: false, onClick: () => {} },
            { label: 'Spelling', checked: false, onClick: () => {} },
            { label: 'Undo', checked: true, onClick: () => {} },
            { label: 'Redo', checked: true, onClick: () => {} },
            { label: 'Sort Ascending', checked: false, onClick: () => {} },
            { label: 'Sort Descending', checked: false, onClick: () => {} },
            'separator',
            { label: 'More Commands...', onClick: () => {} },
            { label: 'Show Below the Ribbon', onClick: () => {} }
          ]}
        />
      )}
    </div>
  )
}

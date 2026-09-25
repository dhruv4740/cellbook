// Name Box + fx + formula input, mirrors the shared ui.edit contract with the grid's in-cell editor.
import { useEffect, useRef, useState } from 'react'
import { useStore } from '../model/store'
import { useUi, setUi, commitEdit, cancelEdit } from './ui'
import { a1, rangeA1 } from '../model/types'
import * as cmds from '../grid/commands'
import { allTables, displayTotalsFormula } from '../model/tables'
import { ChevronDown12Regular, ChevronDown20Regular, ChevronUp20Regular, Dismiss20Regular, Checkmark20Regular, MathFormula20Regular } from '@fluentui/react-icons'
import './FormulaBar.css'

export default function FormulaBar() {
  const store = useStore()
  const ui = useUi()
  const sheet = store.active
  const sel = sheet.selection
  const editing = ui.edit
  const [nameBox, setNameBox] = useState('')
  const [nameEditing, setNameEditing] = useState(false)
  const [expanded, setExpanded] = useState(false)
  const [nameMenu, setNameMenu] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    if (editing && editing.source === 'bar') inputRef.current?.focus()
  }, [editing?.source])

  if (ui.hideFormulaBar) return null

  const rg = sel.ranges[sel.ranges.length - 1]
  const nameBoxText = (() => {
    if (nameEditing) return nameBox
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return a1(rg.r1, rg.c1)
    return rangeA1(rg)
  })()

  const text = editing ? editing.text : (displayTotalsFormula(sheet, sel.active.r, sel.active.c) ?? store.getRaw(sheet, sel.active.r, sel.active.c))
  const tableNames = allTables(store).map((t) => t.table.name)

  const startEditFromBar = (initial?: string) => {
    if (!editing) {
      const raw = store.getRaw(sheet, sel.active.r, sel.active.c)
      setUi({ edit: { r: sel.active.r, c: sel.active.c, text: initial ?? raw, source: 'bar', mode: 'Edit' } })
      store.setMode('Edit')
    }
  }

  return (
    <div className="xl-formulabar">
      <div className="xl-namebox">
        <input
          value={nameBoxText}
          onFocus={() => { setNameEditing(true); setNameBox(nameBoxText) }}
          onChange={(e) => setNameBox(e.target.value)}
          onBlur={() => setNameEditing(false)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { cmds.goTo(nameBox); setNameEditing(false); (e.target as HTMLInputElement).blur() }
            if (e.key === 'Escape') { setNameEditing(false); (e.target as HTMLInputElement).blur() }
          }}
        />
        <button className="xl-namebox-caret" onClick={() => setNameMenu((v) => !v)}><ChevronDown12Regular /></button>
        {nameMenu && (
          <div className="xl-menu-layer" onMouseDown={() => setNameMenu(false)}>
            <div className="xl-menu" style={{ position: 'fixed', left: 4, top: 62, minWidth: 140 }} onMouseDown={(e) => e.stopPropagation()}>
              {tableNames.length === 0 && <div className="xl-menu-item disabled"><span className="xl-menu-label">(No tables)</span></div>}
              {tableNames.map((n) => (
                <div key={n} className="xl-menu-item" onClick={() => { cmds.goTo(n); setNameMenu(false) }}><span className="xl-menu-label">{n}</span></div>
              ))}
            </div>
          </div>
        )}
      </div>
      <div className="xl-fb-splitter">
        <span>⋮</span>
      </div>
      <div className="xl-fb-buttons">
        <button className="xl-fb-btn cancel" disabled={!editing} onClick={() => cancelEdit()} title="Cancel"><Dismiss20Regular /></button>
        <button className="xl-fb-btn enter" disabled={!editing} onClick={() => commitEdit()} title="Enter"><Checkmark20Regular /></button>
        <button className="xl-fb-btn fx" onClick={() => startEditFromBar(text.startsWith('=') ? text : '=')} title="Insert Function"><MathFormula20Regular /></button>
      </div>
      <div className={'xl-fb-input' + (expanded ? ' expanded' : '')}>
        <input
          ref={inputRef}
          value={text}
          onFocus={() => startEditFromBar()}
          onChange={(e) => { if (editing) setUi({ edit: { ...editing, text: e.target.value } }); else startEditFromBar(e.target.value) }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { e.preventDefault(); commitEdit() }
            else if (e.key === 'Escape') { e.preventDefault(); cancelEdit() }
          }}
        />
      </div>
      <button className="xl-fb-expand" onClick={() => setExpanded((v) => !v)} title={expanded ? 'Collapse Formula Bar' : 'Expand Formula Bar'}>
        {expanded ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
      </button>
    </div>
  )
}

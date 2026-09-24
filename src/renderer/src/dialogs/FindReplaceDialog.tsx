import { useState, useRef, useEffect } from 'react'
import { createPortal } from 'react-dom'
import { WorkbookStore } from '../model/store'
import {
  FindOptions,
  FindResult,
  findAll,
  findNext,
  replaceAll
} from './find'
import './FindReplaceDialog.css'

export interface FindReplaceDialogProps {
  initialTab?: 'find' | 'replace'
  store: WorkbookStore
  onClose: () => void
}

export function FindReplaceDialog({
  initialTab = 'find',
  store,
  onClose
}: FindReplaceDialogProps) {
  const [tab, setTab] = useState<'find' | 'replace'>(initialTab)
  const [findText, setFindText] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [findHistory, setFindHistory] = useState<string[]>([])
  const [replaceHistory, setReplaceHistory] = useState<string[]>([])

  const [optionsOpen, setOptionsOpen] = useState(false)
  const [within, setWithin] = useState<'sheet' | 'workbook'>('sheet')
  const [searchBy, setSearchBy] = useState<'rows' | 'columns'>('rows')
  const [lookIn, setLookIn] = useState<'formulas' | 'values' | 'notes' | 'comments'>('formulas')
  const [matchCase, setMatchCase] = useState(false)
  const [matchEntireCell, setMatchEntireCell] = useState(false)

  const [results, setResults] = useState<FindResult[] | null>(null)
  const [selectedIndex, setSelectedIndex] = useState<number | null>(null)
  const [message, setMessage] = useState<string | null>(null)

  // Dragging support
  const [pos, setPos] = useState<{ x: number; y: number } | null>(null)
  const dragRef = useRef<{ dx: number; dy: number } | null>(null)
  const windowRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const handleMove = (e: MouseEvent) => {
      if (dragRef.current) {
        setPos({
          x: e.clientX - dragRef.current.dx,
          y: e.clientY - dragRef.current.dy
        })
      }
    }
    const handleUp = () => {
      dragRef.current = null
    }
    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [])

  const getOptions = (): FindOptions => ({
    query: findText,
    within,
    searchBy,
    lookIn: tab === 'replace' ? 'formulas' : lookIn,
    matchCase,
    matchEntireCell
  })

  const saveHistory = (val: string, isReplace = false) => {
    if (!val) return
    if (isReplace) {
      setReplaceHistory((prev) => (prev.includes(val) ? prev : [val, ...prev.slice(0, 9)]))
    } else {
      setFindHistory((prev) => (prev.includes(val) ? prev : [val, ...prev.slice(0, 9)]))
    }
  }

  const navigateToResult = (res: FindResult) => {
    if (res.sheetIndex !== store.activeIndex) {
      store.setActiveSheet(res.sheetIndex)
    }
    const sheet = store.sheets[res.sheetIndex]
    if (sheet) {
      store.setSelection(
        {
          ranges: [{ r1: res.r, c1: res.c, r2: res.r, c2: res.c }],
          active: { r: res.r, c: res.c }
        },
        sheet
      )
    }
  }

  const handleFindNext = () => {
    if (!findText) return
    saveHistory(findText)
    const active = store.active?.selection?.active || { r: 0, c: 0 }
    const match = findNext(store, getOptions(), {
      sheetIndex: store.activeIndex,
      r: active.r,
      c: active.c
    })
    if (match) {
      navigateToResult(match)
      if (results) {
        const idx = results.findIndex((r) => r.sheetIndex === match.sheetIndex && r.r === match.r && r.c === match.c)
        if (idx >= 0) setSelectedIndex(idx)
      }
    } else {
      setMessage("Microsoft Excel cannot find the data you're searching for. Check your search options, or try a different search term.")
    }
  }

  const handleFindAll = () => {
    if (!findText) return
    saveHistory(findText)
    const matches = findAll(store, getOptions())
    setResults(matches)
    if (matches.length > 0) {
      setSelectedIndex(0)
      navigateToResult(matches[0])
    }
  }

  const handleReplace = () => {
    if (!findText) return
    saveHistory(findText)
    saveHistory(replaceText, true)

    const active = store.active?.selection?.active || { r: 0, c: 0 }
    const sheet = store.active
    const raw = store.getRaw(sheet, active.r, active.c)

    // If current cell matches, replace it
    const options = getOptions()
    const matchCaseOpt = !!options.matchCase
    const matchEntireOpt = !!options.matchEntireCell
    let matchesCurrent = false
    if (matchEntireOpt) {
      matchesCurrent = matchCaseOpt ? raw === findText : raw.toLowerCase() === findText.toLowerCase()
    } else {
      matchesCurrent = matchCaseOpt ? raw.includes(findText) : raw.toLowerCase().includes(findText.toLowerCase())
    }

    if (matchesCurrent) {
      let next = ''
      if (matchEntireOpt) {
        next = replaceText
      } else {
        const re = new RegExp(
          findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
          matchCaseOpt ? '' : 'i'
        )
        next = raw.replace(re, replaceText)
      }
      store.setRaw(sheet, active.r, active.c, next)
    }

    // Always move to next match
    handleFindNext()
  }

  const handleReplaceAll = () => {
    if (!findText) return
    saveHistory(findText)
    saveHistory(replaceText, true)
    const count = replaceAll(store, getOptions(), replaceText)
    setMessage(`All done. We made ${count} replacement${count === 1 ? '' : 's'}.`)
    if (results) {
      setResults(findAll(store, getOptions()))
    }
  }

  return createPortal(
    <div
      ref={windowRef}
      className="xl-fr-window"
      style={pos ? { left: pos.x, top: pos.y, transform: 'none' } : {}}
      onKeyDown={(e) => {
        if (e.key === 'Escape') onClose()
        else if (e.key === 'Enter' && e.target instanceof HTMLInputElement) {
          handleFindNext()
        }
      }}
    >
      {/* Title bar */}
      <div
        className="xl-fr-title"
        onMouseDown={(e) => {
          if (windowRef.current) {
            const r = windowRef.current.getBoundingClientRect()
            dragRef.current = { dx: e.clientX - r.left, dy: e.clientY - r.top }
          }
        }}
      >
        <span>Find and Replace</span>
        <span className="xl-fr-tbtn">?</span>
        <span className="xl-fr-tbtn close" onClick={onClose}>✕</span>
      </div>

      {/* Tabs */}
      <div className="xl-fr-tabs">
        <div
          className={`xl-fr-tab ${tab === 'find' ? 'active' : ''}`}
          onClick={() => setTab('find')}
        >
          Find
        </div>
        <div
          className={`xl-fr-tab ${tab === 'replace' ? 'active' : ''}`}
          onClick={() => setTab('replace')}
        >
          Replace
        </div>
      </div>

      {/* Body */}
      <div className="xl-fr-body">
        <div className="xl-fr-row">
          <span className="xl-fr-label">Fi<u>n</u>d what:</span>
          <div className="xl-fr-combo-wrap">
            <input
              className="xl-fr-input"
              value={findText}
              autoFocus
              onChange={(e) => setFindText(e.target.value)}
              list="xl-find-history"
            />
            <datalist id="xl-find-history">
              {findHistory.map((h, i) => (
                <option key={i} value={h} />
              ))}
            </datalist>
          </div>
          <button className="xl-btn" style={{ minWidth: 65 }} disabled title="No format set">
            Format...
          </button>
        </div>

        {tab === 'replace' && (
          <div className="xl-fr-row">
            <span className="xl-fr-label">Re<u>p</u>lace with:</span>
            <div className="xl-fr-combo-wrap">
              <input
                className="xl-fr-input"
                value={replaceText}
                onChange={(e) => setReplaceText(e.target.value)}
                list="xl-replace-history"
              />
              <datalist id="xl-replace-history">
                {replaceHistory.map((h, i) => (
                  <option key={i} value={h} />
                ))}
              </datalist>
            </div>
            <button className="xl-btn" style={{ minWidth: 65 }} disabled title="No format set">
              Format...
            </button>
          </div>
        )}

        {/* Options section */}
        {optionsOpen && (
          <div className="xl-fr-options-panel">
            <div className="xl-fr-options-grid">
              <div className="xl-fr-row">
                <span style={{ width: 60 }}>Within:</span>
                <select
                  className="xl-fr-select"
                  style={{ flex: 1 }}
                  value={within}
                  onChange={(e) => setWithin(e.target.value as 'sheet' | 'workbook')}
                >
                  <option value="sheet">Sheet</option>
                  <option value="workbook">Workbook</option>
                </select>
              </div>

              <div className="xl-fr-row">
                <label className="xl-fr-check-label">
                  <input
                    type="checkbox"
                    checked={matchCase}
                    onChange={(e) => setMatchCase(e.target.checked)}
                  />
                  <span>Match <u>c</u>ase</span>
                </label>
              </div>

              <div className="xl-fr-row">
                <span style={{ width: 60 }}>Search:</span>
                <select
                  className="xl-fr-select"
                  style={{ flex: 1 }}
                  value={searchBy}
                  onChange={(e) => setSearchBy(e.target.value as 'rows' | 'columns')}
                >
                  <option value="rows">By Rows</option>
                  <option value="columns">By Columns</option>
                </select>
              </div>

              <div className="xl-fr-row">
                <label className="xl-fr-check-label">
                  <input
                    type="checkbox"
                    checked={matchEntireCell}
                    onChange={(e) => setMatchEntireCell(e.target.checked)}
                  />
                  <span>Match entire cell c<u>o</u>ntents</span>
                </label>
              </div>

              <div className="xl-fr-row">
                <span style={{ width: 60 }}>Look in:</span>
                <select
                  className="xl-fr-select"
                  style={{ flex: 1 }}
                  value={tab === 'replace' ? 'formulas' : lookIn}
                  disabled={tab === 'replace'}
                  onChange={(e) => setLookIn(e.target.value as any)}
                >
                  <option value="formulas">Formulas</option>
                  <option value="values">Values</option>
                  <option value="notes">Notes</option>
                  <option value="comments">Comments</option>
                </select>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Buttons */}
      <div className="xl-fr-btn-bar">
        <button
          className="xl-btn"
          style={{ marginRight: 'auto' }}
          onClick={() => setOptionsOpen(!optionsOpen)}
        >
          {optionsOpen ? 'Options <<' : 'Options >>'}
        </button>

        {tab === 'replace' && (
          <>
            <button className="xl-btn" onClick={handleReplaceAll}>
              Replace <u>A</u>ll
            </button>
            <button className="xl-btn" onClick={handleReplace}>
              <u>R</u>eplace
            </button>
          </>
        )}

        <button className="xl-btn" onClick={handleFindAll}>
          Find <u>A</u>ll
        </button>
        <button className="xl-btn primary" onClick={handleFindNext}>
          <u>F</u>ind Next
        </button>
        <button className="xl-btn" onClick={onClose}>
          Close
        </button>
      </div>

      {/* Results table if Find All was clicked */}
      {results !== null && (
        <div className="xl-fr-results-panel">
          <div className="xl-fr-results-table-wrap">
            <table className="xl-fr-results-table">
              <thead>
                <tr>
                  <th style={{ width: 75 }}>Book</th>
                  <th style={{ width: 65 }}>Sheet</th>
                  <th style={{ width: 60 }}>Name</th>
                  <th style={{ width: 55 }}>Cell</th>
                  <th style={{ width: 85 }}>Value</th>
                  <th>Formula</th>
                </tr>
              </thead>
              <tbody>
                {results.map((res, i) => (
                  <tr
                    key={i}
                    className={selectedIndex === i ? 'selected' : ''}
                    onClick={() => {
                      setSelectedIndex(i)
                      navigateToResult(res)
                    }}
                  >
                    <td>{res.book}</td>
                    <td>{res.sheetName}</td>
                    <td></td>
                    <td>{res.cellRef}</td>
                    <td>{res.value}</td>
                    <td>{res.formula}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="xl-fr-results-status">
            {results.length} cell(s) found
          </div>
        </div>
      )}

      {/* Message Modal */}
      {message && (
        <div
          style={{
            position: 'absolute',
            inset: 0,
            background: 'rgba(0,0,0,0.2)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000
          }}
          onClick={() => setMessage(null)}
        >
          <div
            style={{
              background: '#fff',
              border: '1px solid #adadad',
              boxShadow: '0 4px 16px rgba(0,0,0,0.2)',
              width: 320,
              padding: 16
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 16 }}>
              <span style={{ fontSize: 24, color: '#0078d7' }}>ℹ</span>
              <div style={{ fontSize: 12, lineHeight: 1.4 }}>{message}</div>
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <button
                className="xl-btn primary"
                autoFocus
                onClick={() => setMessage(null)}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      )}
    </div>,
    document.body
  )
}

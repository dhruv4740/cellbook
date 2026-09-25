// Excel's AutoFilter column dropdown: Sort commands, Clear Filter, Text/Number Filters submenus,
// a search box and the (Select All)/(Blanks)/value checklist, with OK/Cancel.
import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { store, SheetState } from '../model/store'
import { colName } from '../model/types'
import { distinctColumnValues, BLANK_TOKEN } from '../engine/filter'
import { Dialog } from '../components/Dialog'
import './AutoFilterMenu.css'

interface Props {
  sheet: SheetState
  col: number
  x: number
  y: number
  onClose: () => void
}

type TextOp = 'eq' | 'ne' | 'begins' | 'ends' | 'contains' | 'notContains'
type NumOp = 'eq' | 'ne' | 'gt' | 'lt' | 'between' | 'aboveAvg' | 'belowAvg'

const TEXT_OPS: [TextOp, string][] = [
  ['eq', 'Equals...'], ['ne', 'Does Not Equal...'], ['begins', 'Begins With...'], ['ends', 'Ends With...'],
  ['contains', 'Contains...'], ['notContains', 'Does Not Contain...']
]
const NUM_OPS: [NumOp, string][] = [
  ['eq', 'Equals...'], ['ne', 'Does Not Equal...'], ['gt', 'Greater Than...'], ['lt', 'Less Than...'],
  ['between', 'Between...']
]

export function AutoFilterMenu({ sheet, col, x, y, onClose }: Props): React.ReactElement {
  const ref = useRef<HTMLDivElement>(null)
  const af = sheet.layout.autoFilter!
  const getDisplay = (r: number, c: number) => store.getDisplay(sheet, r, c).text
  const allValues = useMemo(() => distinctColumnValues(af.range, col, getDisplay), [store.version]) // eslint-disable-line react-hooks/exhaustive-deps
  const isNumeric = useMemo(() => {
    for (let r = af.range.r1 + 1; r <= Math.min(af.range.r2, af.range.r1 + 200); r++) {
      const v = store.getValue(sheet, r, col)
      if (v !== null && v !== '' && typeof v !== 'number') return false
    }
    return true
  }, [store.version]) // eslint-disable-line react-hooks/exhaustive-deps

  const current = af.columns[col]
  const [checked, setChecked] = useState<Set<string>>(new Set(current ? current.selected : allValues))
  const [search, setSearch] = useState('')
  const [pos, setPos] = useState({ x, y })
  const [sub, setSub] = useState<'text' | 'number' | null>(null)
  const [customOp, setCustomOp] = useState<{ kind: 'text' | 'number'; op: TextOp | NumOp } | null>(null)

  useEffect(() => {
    const el = ref.current
    if (!el) return
    const r = el.getBoundingClientRect()
    setPos({
      x: x + r.width > window.innerWidth ? Math.max(0, window.innerWidth - r.width - 2) : x,
      y: y + r.height > window.innerHeight ? Math.max(0, window.innerHeight - r.height - 2) : y
    })
  }, [x, y])

  const shown = allValues.filter((v) => v !== BLANK_TOKEN && v.toLowerCase().includes(search.toLowerCase()))
  const hasBlank = allValues.includes(BLANK_TOKEN)
  const allChecked = checked.size === allValues.length
  const someChecked = checked.size > 0 && !allChecked

  function toggleAll(): void {
    setChecked(allChecked ? new Set() : new Set(allValues))
  }
  function toggleOne(v: string): void {
    setChecked((prev) => {
      const next = new Set(prev)
      if (next.has(v)) next.delete(v)
      else next.add(v)
      return next
    })
  }
  function ok(): void {
    store.setColumnFilter(sheet, col, { selected: [...checked] })
    onClose()
  }
  function clearFilter(): void {
    store.setColumnFilter(sheet, col, null)
    onClose()
  }
  function sort(order: 'asc' | 'desc'): void {
    store.sortFilterRange(sheet, col, order)
    onClose()
  }
  function applyCondition(matches: (v: string) => boolean): void {
    store.setColumnFilter(sheet, col, { selected: allValues.filter(matches) })
    onClose()
  }

  const label = getDisplay(af.range.r1, col) || colName(col)

  return createPortal(
    <>
      <div className="xl-afmenu-scrim" onMouseDown={onClose} />
      <div ref={ref} className="xl-afmenu" style={{ left: pos.x, top: pos.y }} onMouseDown={(e) => e.stopPropagation()}>
        <div className="xl-afmenu-item" onClick={() => sort('asc')}>{isNumeric ? 'Sort Smallest to Largest' : 'Sort A to Z'}</div>
        <div className="xl-afmenu-item" onClick={() => sort('desc')}>{isNumeric ? 'Sort Largest to Smallest' : 'Sort Z to A'}</div>
        <div className="xl-afmenu-item disabled">Sort by Color <span className="xl-afmenu-arrow">›</span></div>
        <div className="xl-afmenu-item disabled">Sheet View <span className="xl-afmenu-arrow">›</span></div>
        <div className="xl-afmenu-sep" />
        <div className="xl-afmenu-item" onClick={clearFilter}>Clear Filter From &quot;{label}&quot;</div>
        <div className="xl-afmenu-item disabled">Filter by Color <span className="xl-afmenu-arrow">›</span></div>
        <div className="xl-afmenu-item" onMouseEnter={() => setSub(isNumeric ? 'number' : 'text')} onMouseLeave={() => setSub(null)}>
          {isNumeric ? 'Number Filters' : 'Text Filters'} <span className="xl-afmenu-arrow">›</span>
          {sub && (
            <div className="xl-afmenu-sub">
              {(isNumeric ? NUM_OPS : TEXT_OPS).map(([op, opLabel]) => (
                <div key={op} className="xl-afmenu-item" onClick={() => setCustomOp({ kind: isNumeric ? 'number' : 'text', op })}>{opLabel}</div>
              ))}
              {isNumeric && <div className="xl-afmenu-item" onClick={() => applyCondition((v) => { const n = Number(v); const vals = allValues.map(Number).filter((x) => !isNaN(x)); const avg = vals.reduce((a, b) => a + b, 0) / (vals.length || 1); return n >= avg })}>Above Average</div>}
              {isNumeric && <div className="xl-afmenu-item" onClick={() => applyCondition((v) => { const n = Number(v); const vals = allValues.map(Number).filter((x) => !isNaN(x)); const avg = vals.reduce((a, b) => a + b, 0) / (vals.length || 1); return n < avg })}>Below Average</div>}
              <div className="xl-afmenu-item" onClick={() => setCustomOp({ kind: isNumeric ? 'number' : 'text', op: 'eq' })}>Custom Filter...</div>
            </div>
          )}
        </div>
        <div className="xl-afmenu-search">
          <input placeholder="Search" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        <div className="xl-afmenu-list">
          <label className="xl-afmenu-check"><input type="checkbox" checked={allChecked} ref={(el) => { if (el) el.indeterminate = someChecked }} onChange={toggleAll} /> (Select All)</label>
          {hasBlank && (
            <label className="xl-afmenu-check"><input type="checkbox" checked={checked.has(BLANK_TOKEN)} onChange={() => toggleOne(BLANK_TOKEN)} /> (Blanks)</label>
          )}
          {shown.map((v) => (
            <label key={v} className="xl-afmenu-check"><input type="checkbox" checked={checked.has(v)} onChange={() => toggleOne(v)} /> {v}</label>
          ))}
        </div>
        <div className="xl-afmenu-footer">
          <button className="xl-btn primary" onClick={ok}>OK</button>
          <button className="xl-btn" onClick={onClose}>Cancel</button>
        </div>
      </div>
      {customOp && (
        <CustomFilterDialog
          label={label}
          kind={customOp.kind}
          op={customOp.op}
          onCancel={() => setCustomOp(null)}
          onApply={(matches) => { setCustomOp(null); applyCondition(matches) }}
        />
      )}
    </>,
    document.body
  )
}

function CustomFilterDialog({ label, kind, op, onCancel, onApply }: {
  label: string; kind: 'text' | 'number'; op: TextOp | NumOp
  onCancel: () => void; onApply: (matches: (v: string) => boolean) => void
}): React.ReactElement {
  const [op1, setOp1] = useState<string>(op)
  const [v1, setV1] = useState('')
  const [v2, setV2] = useState('')
  const opsList = kind === 'text' ? TEXT_OPS : NUM_OPS
  const textMatch = (t: TextOp, val: string) => (v: string): boolean => {
    const a = v.toLowerCase(), b = val.toLowerCase()
    switch (t) {
      case 'eq': return a === b
      case 'ne': return a !== b
      case 'begins': return a.startsWith(b)
      case 'ends': return a.endsWith(b)
      case 'contains': return a.includes(b)
      case 'notContains': return !a.includes(b)
    }
  }
  const numMatch = (t: NumOp, val: string, val2: string) => (v: string): boolean => {
    const n = Number(v), a = Number(val), b = Number(val2)
    if (isNaN(n)) return false
    switch (t) {
      case 'eq': return n === a
      case 'ne': return n !== a
      case 'gt': return n > a
      case 'lt': return n < a
      case 'between': return n >= Math.min(a, b) && n <= Math.max(a, b)
      default: return true
    }
  }
  return (
    <Dialog
      title={`Custom AutoFilter — ${label}`}
      width={420}
      onCancel={onCancel}
      onOk={() => onApply(kind === 'text' ? textMatch(op1 as TextOp, v1) : numMatch(op1 as NumOp, v1, v2))}
    >
      <div className="xl-field-row">
        <label>Show rows where:</label>
      </div>
      <div className="xl-afmenu-condrow">
        <select value={op1} onChange={(e) => setOp1(e.target.value)}>
          {opsList.map(([o, l]) => <option key={o} value={o}>{l.replace('...', '')}</option>)}
        </select>
        <input value={v1} onChange={(e) => setV1(e.target.value)} autoFocus />
      </div>
      {kind === 'number' && op1 === 'between' && (
        <div className="xl-afmenu-condrow"><label>and</label><input value={v2} onChange={(e) => setV2(e.target.value)} /></div>
      )}
    </Dialog>
  )
}

// Excel's AutoFilter column dropdown: Sort commands, Clear Filter, Text/Number Filters submenus,
// a search box and the (Select All)/(Blanks)/value checklist, with OK/Cancel.
import { runMaybeBusy } from '../chrome/ui'
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

const ROW_H = 22 // fixed row height keeps the windowed list's maths exact
const LIST_H = 180

type TextOp = 'eq' | 'ne' | 'begins' | 'ends' | 'contains' | 'notContains'
type NumOp = 'eq' | 'ne' | 'gt' | 'lt' | 'between' | 'aboveAvg' | 'belowAvg'
type DateOp = 'eq' | 'before' | 'after' | 'between'

const TEXT_OPS: [TextOp, string][] = [
  ['eq', 'Equals...'], ['ne', 'Does Not Equal...'], ['begins', 'Begins With...'], ['ends', 'Ends With...'],
  ['contains', 'Contains...'], ['notContains', 'Does Not Contain...']
]
const DATE_OPS: [DateOp, string][] = [['eq', 'Equals...'], ['before', 'Before...'], ['after', 'After...'], ['between', 'Between...']]
const isDateFmt = (f: string | undefined): boolean => !!f && /[dmy]/i.test(f.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '')) && !/^general$/i.test(f)
const serialOfDate = (y: number, m: number, d: number): number => Math.round(Date.UTC(y, m, d) / 86400000) + 25569
/** Excel's relative date filters: [first, last] serial of each period around today (weeks start on Sunday). */
function datePeriods(): [string, [number, number]][] {
  const n = new Date(), y = n.getFullYear(), m = n.getMonth(), d = n.getDate()
  const t = serialOfDate(y, m, d), wk = t - n.getDay(), q = Math.floor(m / 3) * 3
  const month = (off: number): [number, number] => [serialOfDate(y, m + off, 1), serialOfDate(y, m + off + 1, 1) - 1]
  const quarter = (off: number): [number, number] => [serialOfDate(y, q + off * 3, 1), serialOfDate(y, q + off * 3 + 3, 1) - 1]
  const year = (off: number): [number, number] => [serialOfDate(y + off, 0, 1), serialOfDate(y + off + 1, 0, 1) - 1]
  return [
    ['Tomorrow', [t + 1, t + 1]], ['Today', [t, t]], ['Yesterday', [t - 1, t - 1]],
    ['Next Week', [wk + 7, wk + 13]], ['This Week', [wk, wk + 6]], ['Last Week', [wk - 7, wk - 1]],
    ['Next Month', month(1)], ['This Month', month(0)], ['Last Month', month(-1)],
    ['Next Quarter', quarter(1)], ['This Quarter', quarter(0)], ['Last Quarter', quarter(-1)],
    ['Next Year', year(1)], ['This Year', year(0)], ['Last Year', year(-1)],
    ['Year to Date', [serialOfDate(y, 0, 1), t]]
  ]
}
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

  const isDate = useMemo(() => {
    for (let r = af.range.r1 + 1; r <= Math.min(af.range.r2, af.range.r1 + 200); r++) {
      const v = store.getValue(sheet, r, col)
      if (v !== null && v !== '') return isDateFmt(store.getStyle(sheet, r, col)?.numFmt)
    }
    return false
  }, [store.version]) // eslint-disable-line react-hooks/exhaustive-deps
  /** Displayed text -> date serial, so a date condition can be tested on the checklist's strings. */
  const serialByText = useMemo(() => {
    const m = new Map<string, number>()
    if (!isDate) return m
    for (let r = af.range.r1 + 1; r <= af.range.r2; r++) { const v = store.getValue(sheet, r, col); if (typeof v === 'number') m.set(getDisplay(r, col), Math.floor(v)) }
    return m
  }, [isDate, store.version]) // eslint-disable-line react-hooks/exhaustive-deps
  const inRange = (lo: number, hi: number) => (v: string): boolean => { const s = serialByText.get(v); return s !== undefined && s >= lo && s <= hi }

  /** Distinct cell fills and font colours used in this column (null = no fill / automatic font). */
  const colors = useMemo(() => {
    const fills = new Set<string | null>(), fonts = new Set<string | null>()
    for (let r = af.range.r1 + 1; r <= af.range.r2; r++) { const st = store.getStyle(sheet, r, col); fills.add(st?.fill ?? null); fonts.add(st?.color ?? null) }
    return { fills: [...fills], fonts: [...fonts] }
  }, [store.version]) // eslint-disable-line react-hooks/exhaustive-deps
  const [colorSub, setColorSub] = useState<'sort' | 'filter' | null>(null)
  function sortByColor(kind: 'fill' | 'font', color: string | null): void {
    onClose()
    runMaybeBusy((af.range.r2 - af.range.r1 + 1) * (af.range.c2 - af.range.c1 + 1), 'Sorting...', () => store.sortFilterRange(sheet, col, 'asc', { kind, color }))
  }
  function filterByColor(kind: 'fill' | 'font', color: string | null): void {
    const texts = new Set<string>()
    for (let r = af.range.r1 + 1; r <= af.range.r2; r++) { const st = store.getStyle(sheet, r, col); if (((kind === 'fill' ? st?.fill : st?.color) ?? null) === color) texts.add(getDisplay(r, col)) }
    applyCondition((v) => texts.has(v))
  }
  const colorMenu = (mode: 'sort' | 'filter'): React.ReactElement => (
    <div className="xl-afmenu-sub">
      {([['fill', 'Cell Color', colors.fills, 'No Fill'], ['font', 'Font Color', colors.fonts, 'Automatic']] as const).map(([kind, title, list, none]) => (
        <div key={kind}>
          <div className="xl-afmenu-item disabled">{mode === 'sort' ? `Sort by ${title}` : `Filter by ${title}`}</div>
          {list.map((c) => (
            <div key={String(c)} className="xl-afmenu-item" onClick={() => (mode === 'sort' ? sortByColor(kind, c) : filterByColor(kind, c))}>
              <span style={{ display: 'inline-block', width: 14, height: 14, marginRight: 6, verticalAlign: 'middle', border: '1px solid #8a8886', background: c ?? (kind === 'fill' ? '#fff' : '#000') }} />{c ?? none}
            </div>
          ))}
        </div>
      ))}
    </div>
  )

  const current = af.columns[col]
  const [checked, setChecked] = useState<Set<string>>(new Set(current ? current.selected : allValues))
  const [search, setSearch] = useState('')
  const [pos, setPos] = useState({ x, y })
  const [sub, setSub] = useState<'text' | 'number' | 'date' | null>(null)
  const [customOp, setCustomOp] = useState<{ kind: 'text' | 'number' | 'date'; op: TextOp | NumOp | DateOp } | null>(null)

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
  const [scrollTop, setScrollTop] = useState(0)
  const headRows = 1 + (allValues.includes(BLANK_TOKEN) ? 1 : 0)
  const firstShown = Math.max(0, Math.floor((scrollTop - headRows * ROW_H) / ROW_H) - 4)
  const lastShown = Math.min(shown.length, firstShown + Math.ceil(LIST_H / ROW_H) + 10)
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
    onClose()
    runMaybeBusy((af.range.r2 - af.range.r1 + 1) * (af.range.c2 - af.range.c1 + 1), 'Sorting...', () => store.sortFilterRange(sheet, col, order))
  }
  function applyCondition(matches: (v: string) => boolean): void {
    store.setColumnFilter(sheet, col, { selected: allValues.filter(matches) })
    onClose()
  }

  const label = getDisplay(af.range.r1, col) || colName(col)

  return createPortal(
    <>
      {!customOp && <div className="xl-afmenu-scrim" onMouseDown={onClose} />}
      <div ref={ref} className="xl-afmenu" style={{ left: pos.x, top: pos.y, visibility: customOp ? 'hidden' : undefined }} onMouseDown={(e) => e.stopPropagation()}>
        <div className="xl-afmenu-item" onClick={() => sort('asc')}>{isNumeric ? 'Sort Smallest to Largest' : 'Sort A to Z'}</div>
        <div className="xl-afmenu-item" onClick={() => sort('desc')}>{isNumeric ? 'Sort Largest to Smallest' : 'Sort Z to A'}</div>
        <div className="xl-afmenu-item" onMouseEnter={() => setColorSub('sort')} onMouseLeave={() => setColorSub(null)}>Sort by Color <span className="xl-afmenu-arrow">›</span>{colorSub === 'sort' && colorMenu('sort')}</div>
        <div className="xl-afmenu-item disabled">Sheet View <span className="xl-afmenu-arrow">›</span></div>
        <div className="xl-afmenu-sep" />
        <div className="xl-afmenu-item" onClick={clearFilter}>Clear Filter From &quot;{label}&quot;</div>
        <div className="xl-afmenu-item" onMouseEnter={() => setColorSub('filter')} onMouseLeave={() => setColorSub(null)}>Filter by Color <span className="xl-afmenu-arrow">›</span>{colorSub === 'filter' && colorMenu('filter')}</div>
        <div className="xl-afmenu-item" onMouseEnter={() => setSub(isDate ? 'date' : isNumeric ? 'number' : 'text')} onMouseLeave={() => setSub(null)}>
          {isDate ? 'Date Filters' : isNumeric ? 'Number Filters' : 'Text Filters'} <span className="xl-afmenu-arrow">›</span>
          {sub && isDate && (
            <div className="xl-afmenu-sub">
              {DATE_OPS.map(([op, opLabel]) => (
                <div key={op} className="xl-afmenu-item" onClick={() => setCustomOp({ kind: 'date', op })}>{opLabel}</div>
              ))}
              <div className="xl-afmenu-sep" />
              {datePeriods().map(([name, [lo, hi]]) => (
                <div key={name} className="xl-afmenu-item" onClick={() => applyCondition(inRange(lo, hi))}>{name}</div>
              ))}
              <div className="xl-afmenu-sep" />
              <div className="xl-afmenu-item" onClick={() => setCustomOp({ kind: 'date', op: 'eq' })}>Custom Filter...</div>
            </div>
          )}
          {sub && !isDate && (
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
        {/* Only the visible slice of the value list is rendered: a column with 50,000 distinct values froze the UI. */}
        <div className="xl-afmenu-list" onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)}>
          <label className="xl-afmenu-check"><input type="checkbox" checked={allChecked} ref={(el) => { if (el) el.indeterminate = someChecked }} onChange={toggleAll} /> (Select All)</label>
          {hasBlank && (
            <label className="xl-afmenu-check"><input type="checkbox" checked={checked.has(BLANK_TOKEN)} onChange={() => toggleOne(BLANK_TOKEN)} /> (Blanks)</label>
          )}
          <div style={{ height: shown.length * ROW_H, position: 'relative' }}>
            {shown.slice(firstShown, lastShown).map((v, i) => (
              <label key={v} className="xl-afmenu-check" style={{ position: 'absolute', top: (firstShown + i) * ROW_H, left: 0, right: 0 }}><input type="checkbox" checked={checked.has(v)} onChange={() => toggleOne(v)} /> {v}</label>
            ))}
          </div>
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
          serialByText={serialByText}
          onCancel={() => setCustomOp(null)}
          onApply={(matches) => { setCustomOp(null); applyCondition(matches) }}
        />
      )}
    </>,
    document.body
  )
}

function CustomFilterDialog({ label, kind, op, serialByText, onCancel, onApply }: {
  label: string; kind: 'text' | 'number' | 'date'; op: TextOp | NumOp | DateOp; serialByText: Map<string, number>
  onCancel: () => void; onApply: (matches: (v: string) => boolean) => void
}): React.ReactElement {
  const [op1, setOp1] = useState<string>(op)
  const [v1, setV1] = useState('')
  const [v2, setV2] = useState('')
  const opsList: [string, string][] = kind === 'text' ? TEXT_OPS : kind === 'date' ? DATE_OPS : NUM_OPS
  const toSerial = (t: string): number => { const d = new Date(t); return isNaN(d.getTime()) ? NaN : serialOfDate(d.getFullYear(), d.getMonth(), d.getDate()) }
  const dateMatch = (t: DateOp, a: number, b: number) => (v: string): boolean => {
    const n = serialByText.get(v)
    if (n === undefined || isNaN(a)) return false
    switch (t) {
      case 'eq': return n === a
      case 'before': return n < a
      case 'after': return n > a
      case 'between': return !isNaN(b) && n >= Math.min(a, b) && n <= Math.max(a, b)
    }
  }
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
      onOk={() => onApply(kind === 'text' ? textMatch(op1 as TextOp, v1) : kind === 'date' ? dateMatch(op1 as DateOp, toSerial(v1), toSerial(v2)) : numMatch(op1 as NumOp, v1, v2))}
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
      {kind !== 'text' && op1 === 'between' && (
        <div className="xl-afmenu-condrow"><label>and</label><input value={v2} onChange={(e) => setV2(e.target.value)} /></div>
      )}
    </Dialog>
  )
}

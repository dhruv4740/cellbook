// Data > Advanced Filter, What-If > Goal Seek, Consolidate.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store, SheetState } from '../model/store'
import { Range, normRange, parseA1, rangeA1, CellValue } from '../model/types'
import { advancedFilterRows, uniqueRows } from '../data/advancedFilter'
import { goalSeek } from '../data/goalSeek'
import { CONSOLIDATE_FUNCTIONS, ConsolidateFn, consolidateByLabels, consolidateByPosition, Cell } from '../data/consolidate'
import { setUi } from '../chrome/ui'

const q = (name: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
const absRef = (sheet: SheetState, r: Range): string => `${q(sheet.name)}!${rangeA1(r).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`

/** "Sheet2!$A$1:$C$5", "'My Sheet'!A1" or "A1:B2" (current sheet) -> sheet + range, or null. */
export function parseRef(text: string): { sheet: SheetState; range: Range } | null {
  let t = text.trim().replace(/^=/, '')
  let sheet = store.active
  const bang = t.lastIndexOf('!')
  if (bang >= 0) {
    let name = t.slice(0, bang)
    if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'")
    const s = store.sheets.find((x) => x.name.toLowerCase() === name.toLowerCase())
    if (!s) return null
    sheet = s
    t = t.slice(bang + 1)
  }
  const range = parseA1(t)
  return range ? { sheet, range: normRange(range) } : null
}

const row = { display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, fontSize: 12 } as const

// ---------- Advanced Filter ----------
export function AdvancedFilterDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const a = sheet.selection.active
  const region = store.currentRegion(sheet, a.r, a.c)
  const [copy, setCopy] = useState(false)
  const [list, setList] = useState(absRef(sheet, region))
  const [crit, setCrit] = useState('')
  const [dest, setDest] = useState('')
  const [unique, setUnique] = useState(false)
  const [err, setErr] = useState('')

  const ok = (): void => {
    const L = parseRef(list)
    if (!L) return setErr('The list range is not valid.')
    const sh = L.sheet, lr = L.range
    let criteria: string[][] = []
    if (crit.trim()) {
      const C = parseRef(crit)
      if (!C) return setErr('The criteria range is not valid.')
      for (let r = C.range.r1; r <= C.range.r2; r++) {
        const rowText: string[] = []
        for (let c = C.range.c1; c <= C.range.c2; c++) rowText.push(store.getRaw(C.sheet, r, c))
        criteria.push(rowText)
      }
    }
    const heads: string[] = []
    for (let c = lr.c1; c <= lr.c2; c++) heads.push(store.getDisplay(sh, lr.r1, c).text)
    const count = lr.r2 - lr.r1
    const read = (i: number, c: number) => ({ text: store.getDisplay(sh, lr.r1 + 1 + i, lr.c1 + c).text, value: store.getValue(sh, lr.r1 + 1 + i, lr.c1 + c) as CellValue })
    let rows = advancedFilterRows(heads, count, read, criteria)
    if (unique) rows = uniqueRows(rows, (i) => heads.map((_h, c) => read(i, c).text).join('\u0001'))
    if (copy) {
      const D = parseRef(dest)
      if (!D) return setErr('The Copy to range is not valid.')
      const cells: { r: number; c: number; raw: string }[] = []
      const put = (i: number, srcRow: number): void => {
        for (let c = lr.c1; c <= lr.c2; c++) {
          const raw = store.getRaw(sh, srcRow, c)
          const v = store.getValue(sh, srcRow, c)
          cells.push({ r: D.range.r1 + i, c: D.range.c1 + (c - lr.c1), raw: raw.startsWith('=') ? (v === null ? '' : typeof v === 'object' ? String((v as { error: string }).error) : String(v)) : raw })
        }
      }
      put(0, lr.r1)
      rows.forEach((i, k) => put(k + 1, lr.r1 + 1 + i))
      onClose()
      store.setCells(D.sheet, cells, 'Advanced Filter')
      return
    }
    const keep = new Set(rows.map((i) => lr.r1 + 1 + i))
    onClose()
    store.updateLayout(sh, 'Advanced Filter', (l) => {
      if (l.advFilter) for (const r of l.advFilter.hidden) delete l.hiddenRows[r]
      const hidden: number[] = []
      for (let r = lr.r1 + 1; r <= lr.r2; r++) if (!keep.has(r) && !l.hiddenRows[r]) { l.hiddenRows[r] = true; hidden.push(r) }
      l.advFilter = { hidden }
    })
  }

  return (
    <Dialog title="Advanced Filter" width={380} onOk={ok} onCancel={onClose}>
      <div style={{ fontSize: 12 }}>
        {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
        <div style={{ marginBottom: 4 }}>Action</div>
        <label style={{ display: 'flex', gap: 6 }}><input type="radio" checked={!copy} onChange={() => setCopy(false)} /> Filter the list, in-place</label>
        <label style={{ display: 'flex', gap: 6, marginBottom: 8 }}><input type="radio" checked={copy} onChange={() => setCopy(true)} /> Copy to another location</label>
        <div style={row}><span style={{ width: 90 }}>List range:</span><input style={{ flex: 1 }} value={list} onChange={(e) => { setList(e.target.value); setErr('') }} /></div>
        <div style={row}><span style={{ width: 90 }}>Criteria range:</span><input style={{ flex: 1 }} value={crit} onChange={(e) => { setCrit(e.target.value); setErr('') }} /></div>
        <div style={row}><span style={{ width: 90, opacity: copy ? 1 : 0.5 }}>Copy to:</span><input style={{ flex: 1 }} disabled={!copy} value={dest} onChange={(e) => { setDest(e.target.value); setErr('') }} /></div>
        <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={unique} onChange={(e) => setUnique(e.target.checked)} /> Unique records only</label>
      </div>
    </Dialog>
  )
}

// ---------- Goal Seek ----------
export function GoalSeekDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const a = sheet.selection.active
  const [setCell, setSetCell] = useState(`${q(sheet.name)}!$${String.fromCharCode(65 + Math.min(a.c, 25))}$${a.r + 1}`)
  const [to, setTo] = useState('')
  const [by, setBy] = useState('')
  const [err, setErr] = useState('')

  const ok = (): void => {
    const S = parseRef(setCell), B = parseRef(by)
    const target = Number(to)
    if (!S || S.range.r1 !== S.range.r2 || S.range.c1 !== S.range.c2) return setErr('Reference must be to a single cell.')
    if (!store.getRaw(S.sheet, S.range.r1, S.range.c1).startsWith('=')) return setErr('Cell must contain a formula.')
    if (to.trim() === '' || isNaN(target)) return setErr('Target value must be a number.')
    if (!B || B.range.r1 !== B.range.r2 || B.range.c1 !== B.range.c2) return setErr('Reference must be to a single cell.')
    const start = store.getValue(B.sheet, B.range.r1, B.range.c1)
    const startNum = typeof start === 'number' ? start : 0
    let result: ReturnType<typeof goalSeek> | null = null
    store.transact('Goal Seek', () => {
      result = goalSeek((x) => {
        store.setRaw(B.sheet, B.range.r1, B.range.c1, String(x))
        if (store.calcMode === 'manual') store.recalculate()
        const v = store.getValue(S.sheet, S.range.r1, S.range.c1)
        return typeof v === 'number' ? v : null
      }, target, startNum)
      store.setRaw(B.sheet, B.range.r1, B.range.c1, String(result.found ? result.x : startNum))
      if (store.calcMode === 'manual') store.recalculate()
    })
    const res = result as unknown as ReturnType<typeof goalSeek>
    const cur = store.getDisplay(S.sheet, S.range.r1, S.range.c1).text
    const setName = rangeA1(S.range)
    onClose()
    setUi({
      dialog: {
        kind: 'goalSeekStatus',
        props: {
          text: res.found ? `Goal Seeking with Cell ${setName} found a solution.` : `Goal Seeking with Cell ${setName} may not have found a solution.`,
          target: String(target), current: cur,
          onCancel: () => store.undo()
        }
      }
    })
  }

  return (
    <Dialog title="Goal Seek" width={340} onOk={ok} onCancel={onClose}>
      <div style={{ fontSize: 12 }}>
        {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
        <div style={row}><span style={{ width: 100 }}>Set cell:</span><input style={{ flex: 1 }} value={setCell} onChange={(e) => { setSetCell(e.target.value); setErr('') }} /></div>
        <div style={row}><span style={{ width: 100 }}>To value:</span><input style={{ flex: 1 }} value={to} onChange={(e) => { setTo(e.target.value); setErr('') }} autoFocus /></div>
        <div style={row}><span style={{ width: 100 }}>By changing cell:</span><input style={{ flex: 1 }} value={by} onChange={(e) => { setBy(e.target.value); setErr('') }} /></div>
      </div>
    </Dialog>
  )
}

export function GoalSeekStatusDialog({ text, target, current, onOk, onCancel }: { text: string; target: string; current: string; onOk: () => void; onCancel: () => void }) {
  return (
    <Dialog title="Goal Seek Status" width={340} onOk={onOk} onCancel={onCancel}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 8 }}>{text}</div>
        <div>Target value: {target}</div>
        <div>Current value: {current}</div>
      </div>
    </Dialog>
  )
}

// ---------- Consolidate ----------
export function ConsolidateDialog({ onClose }: { onClose: () => void }) {
  const [fn, setFn] = useState<ConsolidateFn>('Sum')
  const [refs, setRefs] = useState<string[]>([])
  const [text, setText] = useState('')
  const [top, setTop] = useState(false)
  const [left, setLeft] = useState(false)
  const [err, setErr] = useState('')

  const add = (): void => {
    const p = parseRef(text)
    if (!p) return setErr('Reference is not valid.')
    setRefs((r) => [...r, absRef(p.sheet, p.range)])
    setText(''); setErr('')
  }
  const ok = (): void => {
    const list = text.trim() ? [...refs, (() => { const p = parseRef(text); return p ? absRef(p.sheet, p.range) : '' })()].filter(Boolean) : refs
    if (!list.length) return setErr('Reference is not valid.')
    const sources = list.map((s) => {
      const p = parseRef(s)!
      const data: Cell[][] = []
      for (let r = p.range.r1; r <= p.range.r2; r++) {
        const rowVals: Cell[] = []
        for (let c = p.range.c1; c <= p.range.c2; c++) { const v = store.getValue(p.sheet, r, c); rowVals.push(typeof v === 'object' && v !== null ? null : (v as Cell)) }
        data.push(rowVals)
      }
      return { data }
    })
    const sheet = store.active
    const at = sheet.selection.active
    const out: Cell[][] = top || left ? consolidateByLabels(sources, fn, top, left) : consolidateByPosition(sources, fn)
    const cells: { r: number; c: number; raw: string }[] = []
    out.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== '') cells.push({ r: at.r + i, c: at.c + j, raw: typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v) }) }))
    onClose()
    if (cells.length) store.setCells(sheet, cells, 'Consolidate')
  }

  return (
    <Dialog title="Consolidate" width={420} onOk={ok} onCancel={onClose}>
      <div style={{ fontSize: 12 }}>
        {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
        <div style={row}><span style={{ width: 70 }}>Function:</span>
          <select value={fn} onChange={(e) => setFn(e.target.value as ConsolidateFn)}>{CONSOLIDATE_FUNCTIONS.map((f) => <option key={f}>{f}</option>)}</select></div>
        <div>Reference:</div>
        <div style={row}><input style={{ flex: 1 }} value={text} onChange={(e) => { setText(e.target.value); setErr('') }} /><button className="xl-btn" onClick={add}>Add</button></div>
        <div>All references:</div>
        <div style={{ border: '1px solid #c8c8c8', height: 90, overflow: 'auto', marginBottom: 6 }}>
          {refs.map((r, i) => <div key={i} style={{ padding: '1px 4px', display: 'flex', justifyContent: 'space-between' }}>{r}<span style={{ cursor: 'pointer' }} onClick={() => setRefs(refs.filter((_x, k) => k !== i))}>✕</span></div>)}
        </div>
        <div>Use labels in</div>
        <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={top} onChange={(e) => setTop(e.target.checked)} /> Top row</label>
        <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={left} onChange={(e) => setLeft(e.target.checked)} /> Left column</label>
      </div>
    </Dialog>
  )
}

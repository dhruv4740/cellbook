// Insert > PivotTable: source range, location, and the field list (Rows / Columns / Values).
import { useMemo, useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { normRange, colName, parseA1 } from '../model/types'
import { PivotDef, PivotFunc, PIVOT_FUNC_LABEL, renderPivot } from '../data/pivot'

type Area = 'none' | 'rows' | 'cols' | 'values'

function quote(n: string): string { return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'` }

export function PivotDialog({ onClose, edit }: { onClose: () => void; edit?: { sheetUid: number; id: string } }) {
  const editSheet = edit ? store.sheetByUid(edit.sheetUid) : undefined
  const existing = editSheet?.layout.pivots?.find((p) => p.id === edit!.id)
  const sheet = store.active
  const initialRange = useMemo(() => {
    if (existing) return existing.source
    const n = normRange(sheet.selection.ranges[sheet.selection.ranges.length - 1])
    return n.r1 !== n.r2 || n.c1 !== n.c2 ? store.clampToUsed(sheet, n) : store.currentRegion(sheet, sheet.selection.active.r, sheet.selection.active.c)
  }, [])
  const srcSheetName = existing?.sourceSheet ?? sheet.name
  const [rangeText, setRangeText] = useState(`${quote(srcSheetName)}!$${colName(initialRange.c1)}$${initialRange.r1 + 1}:$${colName(initialRange.c2)}$${initialRange.r2 + 1}`)
  const [where, setWhere] = useState<'new' | 'existing'>('new')
  const [destText, setDestText] = useState('')
  const parsed = useMemo(() => {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(rangeText.trim())
    const sn = m ? (m[1] ?? m[2]).replace(/''/g, "'") : srcSheetName
    const rg = parseA1(m ? m[3] : rangeText)
    const sh = store.sheets.find((s) => s.name.toLowerCase() === sn.toLowerCase())
    return sh && rg ? { sh, rg } : null
  }, [rangeText])
  const fields = parsed ? Array.from({ length: parsed.rg.c2 - parsed.rg.c1 + 1 }, (_, i) => parsed.rg.c1 + i) : []
  const head = (c: number) => (parsed ? store.getDisplay(parsed.sh, parsed.rg.r1, c).text : '') || `Column${c + 1}`
  const [areas, setAreas] = useState<Record<number, Area>>(() => {
    const a: Record<number, Area> = {}
    if (existing) {
      existing.rows.forEach((c) => (a[c] = 'rows'))
      existing.cols.forEach((c) => (a[c] = 'cols'))
      existing.values.forEach((v) => (a[v.col] = 'values'))
    }
    return a
  })
  const [funcs, setFuncs] = useState<Record<number, PivotFunc>>(() => Object.fromEntries((existing?.values ?? []).map((v) => [v.col, v.func])))
  const [order, setOrder] = useState<number[]>(() => existing ? [...existing.rows, ...existing.cols, ...existing.values.map((v) => v.col)] : [])
  const [err, setErr] = useState<string | null>(null)

  const setArea = (c: number, a: Area) => {
    setAreas((x) => ({ ...x, [c]: a }))
    setOrder((o) => [...o.filter((k) => k !== c), c])
    if (a === 'values' && !funcs[c] && parsed) {
      const numeric = typeof store.getValue(parsed.sh, parsed.rg.r1 + 1, c) === 'number'
      setFuncs((f) => ({ ...f, [c]: numeric ? 'sum' : 'count' }))
    }
  }
  const inArea = (a: Area) => order.filter((c) => areas[c] === a)

  const ok = () => {
    if (!parsed) { setErr('The data source reference is not valid.'); return }
    if (parsed.rg.r2 <= parsed.rg.r1) { setErr('The PivotTable data source needs a header row and data.'); return }
    const rows = inArea('rows'), cols = inArea('cols'), vals = inArea('values')
    if (!rows.length && !vals.length) { setErr('Put at least one field in Rows or Values.'); return }
    if (rows.length > 2) { setErr('Cellbook supports up to two Rows fields.'); return }
    if (cols.length > 1) { setErr('Cellbook supports one Columns field.'); return }
    if (cols.length && vals.length > 1) { setErr('With a Columns field, use one Values field.'); return }
    let dest = editSheet, anchor = existing?.anchor ?? { r: 2, c: 0 }
    store.transact('PivotTable', () => {
      if (!existing) {
        if (where === 'new') {
          let n = 1
          while (store.sheets.some((s) => s.name.toLowerCase() === `sheet${n}`.toLowerCase())) n++
          store.addSheet(`Sheet${n}`, store.activeIndex)
          dest = store.active
        } else {
          const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(destText.trim())
          const sn = m ? (m[1] ?? m[2]).replace(/''/g, "'") : store.active.name
          const rg = parseA1(m ? m[3] : destText)
          dest = store.sheets.find((s) => s.name.toLowerCase() === sn.toLowerCase())
          if (!dest || !rg) return
          anchor = { r: rg.r1, c: rg.c1 }
        }
      }
      const def: PivotDef = {
        id: existing?.id ?? 'pvt' + Date.now().toString(36),
        sourceSheet: parsed.sh.name, source: parsed.rg, anchor, rows, cols,
        values: vals.map((c) => ({ col: c, func: funcs[c] ?? 'sum' })),
        lastSize: existing?.lastSize
      }
      const e = renderPivot(store, dest!, def)
      if (e) setErr(e)
    })
    if (!dest) { setErr('The destination reference is not valid.'); return }
    const di = store.sheets.indexOf(dest)
    if (di >= 0) {
      store.setActiveSheet(di)
      store.setSelection({ ranges: [{ r1: anchor.r, c1: anchor.c, r2: anchor.r, c2: anchor.c }], active: anchor }, dest)
    }
    onClose()
  }

  const areaLabel: Record<Area, string> = { none: '—', rows: 'Rows', cols: 'Columns', values: 'Values' }
  return (
    <Dialog title={existing ? 'PivotTable Fields' : 'Create PivotTable'} width={520} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <div className="xl-field-row"><label style={{ width: 110 }}>Table/Range:</label><input value={rangeText} onChange={(e) => setRangeText(e.target.value)} style={{ flex: 1 }} disabled={!!existing} /></div>
        {!existing && (
          <div style={{ marginTop: 6 }}>
            <label><input type="radio" checked={where === 'new'} onChange={() => setWhere('new')} /> New Worksheet</label>{'  '}
            <label><input type="radio" checked={where === 'existing'} onChange={() => setWhere('existing')} /> Existing Worksheet:</label>{' '}
            <input value={destText} onChange={(e) => setDestText(e.target.value)} disabled={where !== 'existing'} placeholder="Sheet2!A3" style={{ width: 120 }} />
          </div>
        )}
        <div style={{ marginTop: 10, marginBottom: 4 }}>Choose fields to add to report:</div>
        <div style={{ border: '1px solid #c8c8c8', maxHeight: 220, overflow: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              {fields.map((c) => (
                <tr key={c} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '2px 6px' }}>{head(c)}</td>
                  <td style={{ width: 110 }}>
                    <select value={areas[c] ?? 'none'} onChange={(e) => setArea(c, e.target.value as Area)} style={{ width: '100%' }}>
                      {(['none', 'rows', 'cols', 'values'] as Area[]).map((a) => <option key={a} value={a}>{areaLabel[a]}</option>)}
                    </select>
                  </td>
                  <td style={{ width: 100 }}>
                    {areas[c] === 'values' && (
                      <select value={funcs[c] ?? 'sum'} onChange={(e) => setFuncs((f) => ({ ...f, [c]: e.target.value as PivotFunc }))} style={{ width: '100%' }}>
                        {(Object.keys(PIVOT_FUNC_LABEL) as PivotFunc[]).map((f) => <option key={f} value={f}>{PIVOT_FUNC_LABEL[f]}</option>)}
                      </select>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ color: '#666', marginTop: 6 }}>
          Rows: {inArea('rows').map(head).join(', ') || '—'} · Columns: {inArea('cols').map(head).join(', ') || '—'} · Values: {inArea('values').map((c) => `${PIVOT_FUNC_LABEL[funcs[c] ?? 'sum']} of ${head(c)}`).join(', ') || '—'}
        </div>
        <div style={{ color: '#666', marginTop: 4 }}>After changing the source data, use Data &gt; Refresh All (or right-click the PivotTable &gt; Refresh). The .xlsx file keeps the PivotTable's values.</div>
        {err && <div style={{ color: '#a4262c', marginTop: 6 }}>{err}</div>}
      </div>
    </Dialog>
  )
}

// Page Layout > Page Setup: Page / Margins / Header/Footer / Sheet tabs, plus the Custom Header/Footer editor.
import { useRef, useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { PageSetup } from '../model/types'
import { PAPER_SIZES, paperOf, marginsOf, fitOf, splitHf, joinHf, hfPresets, hfPlain, HfParts, HfSection, HfContext } from '../model/pageSetup'
import { PageSetupTab } from '../grid/pageSetupCommands'
import { DEFAULT_AUTHOR } from '../model/defaults'
import './FormatCellsDialog.css'

const TABS: PageSetupTab[] = ['Page', 'Margins', 'Header/Footer', 'Sheet']
const row = { display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 } as const
const num = { width: 60 } as const

function quote(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}
function colName(c: number): string { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26) } return s }
const colIndex = (t: string): number => [...t.toUpperCase()].reduce((v, ch) => v * 26 + ch.charCodeAt(0) - 64, 0) - 1

const titleText = (t: [number, number] | undefined, kind: 'r' | 'c'): string =>
  !t ? '' : kind === 'r' ? `$${t[0] + 1}:$${t[1] + 1}` : `$${colName(t[0])}:$${colName(t[1])}`
function parseSpan(text: string, kind: 'r' | 'c'): [number, number] | null | undefined {
  const t = text.trim().replace(/^.*!/, '')
  if (!t) return null
  const m = kind === 'r' ? /^\$?(\d+):\$?(\d+)$/.exec(t) : /^\$?([A-Za-z]+):\$?([A-Za-z]+)$/.exec(t)
  if (!m) return undefined
  const a = kind === 'r' ? +m[1] - 1 : colIndex(m[1]), b = kind === 'r' ? +m[2] - 1 : colIndex(m[2])
  return [Math.min(a, b), Math.max(a, b)]
}
function areaText(): string {
  const n = store.names.find((x) => /^_xlnm\.Print_Area$/i.test(x.name) && x.sheetUid === store.active.uid)
  return n ? n.ref.replace(/^=/, '').replace(/^.*!/, '') : ''
}
function parseArea(text: string): string | null | undefined {
  const t = text.trim().replace(/^.*!/, '').replace(/\$/g, '').toUpperCase()
  if (!t) return null
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(t)
  if (!m) return undefined
  const a = `$${m[1]}$${m[2]}`
  return `${quote(store.active.name)}!${m[3] ? `${a}:$${m[3]}$${m[4]}` : a}`
}

// "&P" in the file is shown as "&[Page]" in Excel's editor.
const TOKENS: [string, string][] = [['P', 'Page'], ['N', 'Pages'], ['D', 'Date'], ['T', 'Time'], ['Z', 'Path'], ['F', 'File'], ['A', 'Tab'], ['G', 'Picture']]
export function hfToDisplay(code: string): string {
  let out = ''
  for (let i = 0; i < code.length; i++) {
    if (code[i] === '&' && i + 1 < code.length) {
      const t = TOKENS.find((x) => x[0] === code[i + 1])
      out += t ? `&[${t[1]}]` : code[i] + code[i + 1]
      i++
    } else out += code[i]
  }
  return out
}
export function hfFromDisplay(text: string): string {
  let out = ''
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '&' && text[i + 1] === '[') {
      const end = text.indexOf(']', i)
      const t = end > 0 ? TOKENS.find((x) => x[1] === text.slice(i + 2, end)) : undefined
      if (t) { out += '&' + t[0]; i = end; continue }
    }
    if (text[i] === '&' && i + 1 < text.length) { out += text[i] + text[i + 1]; i++; continue }
    out += text[i]
  }
  return out
}

export function PageSetupDialog({ initialTab = 'Page', onClose }: { initialTab?: PageSetupTab; onClose: () => void }) {
  const sheet = store.active
  const [tab, setTab] = useState<PageSetupTab>(initialTab)
  const [ps, setPs] = useState<PageSetup>(() => ({ ...(sheet.layout.pageSetup ?? {}), hf: { ...(sheet.layout.pageSetup?.hf ?? {}) } }))
  const set = (p: Partial<PageSetup>) => setPs((cur) => ({ ...cur, ...p }))
  const setHf = (p: NonNullable<PageSetup['hf']>) => setPs((cur) => ({ ...cur, hf: { ...cur.hf, ...p } }))
  const titles = store.printTitles(sheet)
  const [initial] = useState(() => ({ area: areaText(), rows: titleText(titles.rows, 'r'), cols: titleText(titles.cols, 'c') }))
  const [area, setArea] = useState(initial.area)
  const [rowsText, setRowsText] = useState(initial.rows)
  const [colsText, setColsText] = useState(initial.cols)
  const [custom, setCustom] = useState<'Header' | 'Footer' | null>(null)
  const [err, setErr] = useState('')
  const m = marginsOf(ps)
  const fit = fitOf(ps)
  const hf = ps.hf ?? {}
  const ctx: HfContext = { page: 1, pages: 1, sheet: sheet.name, file: store.fileName || 'Book1', path: '', now: new Date() }

  const ok = (): void => {
    const a = parseArea(area), r = parseSpan(rowsText, 'r'), c = parseSpan(colsText, 'c')
    if (a === undefined) { setTab('Sheet'); setErr('The print area reference is not valid.'); return }
    if (r === undefined || c === undefined) { setTab('Sheet'); setErr('The print titles reference is not valid.'); return }
    onClose()
    store.updateLayout(sheet, 'Page Setup', (l) => { l.pageSetup = ps })
    if (area !== initial.area) store.setPrintArea(sheet, a)
    if (rowsText !== initial.rows || colsText !== initial.cols) store.setPrintTitles(sheet, { rows: r, cols: c })
  }

  const presetSelect = (kind: 'Header' | 'Footer') => {
    const key = `odd${kind}` as 'oddHeader' | 'oddFooter'
    const presets = hfPresets({ user: DEFAULT_AUTHOR })
    const cur = hf[key] ?? ''
    const known = presets.find((p) => p.code === cur)
    return (
      <>
        <div style={{ margin: '8px 0 3px' }}>{kind}:</div>
        <select style={{ width: '100%' }} value={known ? cur : '__custom'} onChange={(e) => setHf({ [key]: e.target.value === '__custom' ? cur : e.target.value })}>
          {presets.map((p) => <option key={p.label} value={p.code}>{p.label}</option>)}
          {!known && <option value="__custom">{hfPlain(cur.replace(/&[LCR]/g, ' '), ctx).trim() || '(custom)'}</option>}
        </select>
        <div style={{ border: '1px solid #999', background: '#fff', minHeight: 44, margin: '6px 0', display: 'flex', justifyContent: 'space-between', padding: 4, fontSize: 12 }}>
          {(['left', 'center', 'right'] as HfSection[]).map((s) => <span key={s} style={{ flex: 1, textAlign: s }}>{hfPlain(splitHf(cur)[s], ctx)}</span>)}
        </div>
      </>
    )
  }

  if (custom) {
    return <CustomHfDialog kind={custom} hf={hf} onOk={(p) => { setHf(p); setCustom(null) }} onCancel={() => setCustom(null)} />
  }

  return (
    <Dialog title="Page Setup" width={480} onOk={ok} onCancel={onClose}>
      <div className="xl-fc-tabs">
        {TABS.map((t) => <div key={t} className={`xl-fc-tab ${tab === t ? 'active' : ''}`} onClick={() => setTab(t)}>{t}</div>)}
      </div>
      <div className="xl-fc-pane" style={{ fontSize: 12, minHeight: 300 }}>
        {tab === 'Page' && (
          <>
            <fieldset style={{ marginBottom: 8 }}><legend>Orientation</legend>
              <label style={row}><input type="radio" checked={!ps.landscape} onChange={() => set({ landscape: false })} /> Portrait
                <input type="radio" checked={!!ps.landscape} onChange={() => set({ landscape: true })} style={{ marginLeft: 24 }} /> Landscape</label>
            </fieldset>
            <fieldset style={{ marginBottom: 8 }}><legend>Scaling</legend>
              <label style={row}><input type="radio" checked={!fit.on} onChange={() => set({ fitToPage: false })} /> Adjust to:
                <input type="number" style={num} min={10} max={400} value={ps.scale ?? 100} disabled={fit.on} onChange={(e) => set({ scale: Math.min(400, Math.max(10, +e.target.value || 100)) })} /> % normal size</label>
              <label style={row}><input type="radio" checked={fit.on} onChange={() => set({ fitToPage: true, fitToWidth: fit.w, fitToHeight: fit.h })} /> Fit to:
                <input type="number" style={num} min={0} max={32767} value={fit.w} disabled={!fit.on} onChange={(e) => set({ fitToWidth: Math.max(0, +e.target.value || 0) })} /> page(s) wide by
                <input type="number" style={num} min={0} max={32767} value={fit.h} disabled={!fit.on} onChange={(e) => set({ fitToHeight: Math.max(0, +e.target.value || 0) })} /> tall</label>
            </fieldset>
            <div style={row}><span style={{ width: 80 }}>Paper size:</span>
              <select style={{ flex: 1 }} value={paperOf(ps).code} onChange={(e) => set({ paperSize: +e.target.value })}>
                {PAPER_SIZES.map((p) => <option key={p.code} value={p.code}>{p.name}</option>)}
              </select></div>
            <div style={row}><span style={{ width: 80 }}>First page number:</span>
              <input style={num} placeholder="Auto" value={ps.firstPageNumber ?? ''} onChange={(e) => { const v = parseInt(e.target.value, 10); set({ firstPageNumber: Number.isFinite(v) ? v : undefined }) }} /></div>
          </>
        )}
        {tab === 'Margins' && (
          <>
            {(['top', 'bottom', 'left', 'right', 'header', 'footer'] as const).map((k) => (
              <div key={k} style={row}><span style={{ width: 70 }}>{k[0].toUpperCase() + k.slice(1)}:</span>
                <input type="number" step={0.05} min={0} style={num} value={m[k]} onChange={(e) => set({ margins: { ...m, [k]: Math.max(0, +e.target.value || 0) } })} /> in</div>
            ))}
            <fieldset><legend>Center on page</legend>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!ps.centerH} onChange={(e) => set({ centerH: e.target.checked })} /> Horizontally</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!ps.centerV} onChange={(e) => set({ centerV: e.target.checked })} /> Vertically</label>
            </fieldset>
          </>
        )}
        {tab === 'Header/Footer' && (
          <>
            {presetSelect('Header')}
            <button className="xl-btn" onClick={() => setCustom('Header')}>Custom Header...</button>
            {presetSelect('Footer')}
            <button className="xl-btn" onClick={() => setCustom('Footer')}>Custom Footer...</button>
            <div style={{ marginTop: 10 }}>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!hf.differentOddEven} onChange={(e) => setHf({ differentOddEven: e.target.checked })} /> Different odd and even pages</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!hf.differentFirst} onChange={(e) => setHf({ differentFirst: e.target.checked })} /> Different first page</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={hf.scaleWithDoc !== false} onChange={(e) => setHf({ scaleWithDoc: e.target.checked })} /> Scale with document</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={hf.alignWithMargins !== false} onChange={(e) => setHf({ alignWithMargins: e.target.checked })} /> Align with page margins</label>
            </div>
          </>
        )}
        {tab === 'Sheet' && (
          <>
            {err && <div style={{ color: '#a4262c', marginBottom: 6 }}>{err}</div>}
            <div style={row}><span style={{ width: 130 }}>Print area:</span><input style={{ flex: 1 }} value={area} onChange={(e) => { setArea(e.target.value); setErr('') }} /></div>
            <div style={row}><span style={{ width: 130 }}>Rows to repeat at top:</span><input style={{ flex: 1 }} value={rowsText} onChange={(e) => { setRowsText(e.target.value); setErr('') }} /></div>
            <div style={row}><span style={{ width: 130 }}>Columns to repeat at left:</span><input style={{ flex: 1 }} value={colsText} onChange={(e) => { setColsText(e.target.value); setErr('') }} /></div>
            <fieldset style={{ marginBottom: 8 }}><legend>Print</legend>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!ps.gridlines} onChange={(e) => set({ gridlines: e.target.checked })} /> Gridlines</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!ps.blackAndWhite} onChange={(e) => set({ blackAndWhite: e.target.checked })} /> Black and white</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!ps.draft} onChange={(e) => set({ draft: e.target.checked })} /> Draft quality</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="checkbox" checked={!!ps.headings} onChange={(e) => set({ headings: e.target.checked })} /> Row and column headings</label>
              <div style={row}><span style={{ width: 90 }}>Comments:</span>
                <select value={ps.comments ?? ''} onChange={(e) => set({ comments: (e.target.value || undefined) as PageSetup['comments'] })}>
                  <option value="">(None)</option><option value="atEnd">At end of sheet</option><option value="asDisplayed">As displayed on sheet</option>
                </select></div>
              <div style={row}><span style={{ width: 90 }}>Cell errors as:</span>
                <select value={ps.errors ?? ''} onChange={(e) => set({ errors: (e.target.value || undefined) as PageSetup['errors'] })}>
                  <option value="">displayed</option><option value="blank">&lt;blank&gt;</option><option value="dash">--</option><option value="NA">#N/A</option>
                </select></div>
            </fieldset>
            <fieldset><legend>Page order</legend>
              <label style={{ display: 'flex', gap: 6 }}><input type="radio" checked={ps.pageOrder !== 'overThenDown'} onChange={() => set({ pageOrder: undefined })} /> Down, then over</label>
              <label style={{ display: 'flex', gap: 6 }}><input type="radio" checked={ps.pageOrder === 'overThenDown'} onChange={() => set({ pageOrder: 'overThenDown' })} /> Over, then down</label>
            </fieldset>
          </>
        )}
      </div>
    </Dialog>
  )
}

const INSERTS: { label: string; text: string }[] = [
  { label: 'Page Number', text: '&[Page]' }, { label: 'Number of Pages', text: '&[Pages]' }, { label: 'Date', text: '&[Date]' },
  { label: 'Time', text: '&[Time]' }, { label: 'File Path', text: '&[Path]' }, { label: 'File Name', text: '&[File]' }, { label: 'Sheet Name', text: '&[Tab]' }
]

function CustomHfDialog({ kind, hf, onOk, onCancel }: { kind: 'Header' | 'Footer'; hf: NonNullable<PageSetup['hf']>; onOk: (p: NonNullable<PageSetup['hf']>) => void; onCancel: () => void }) {
  const variants = [
    { key: `odd${kind}`, label: hf.differentOddEven ? `Odd Page ${kind}` : kind },
    ...(hf.differentOddEven ? [{ key: `even${kind}`, label: `Even Page ${kind}` }] : []),
    ...(hf.differentFirst ? [{ key: `first${kind}`, label: `First Page ${kind}` }] : [])
  ]
  const [vi, setVi] = useState(0)
  const [parts, setParts] = useState<Record<string, HfParts>>(() => {
    const o: Record<string, HfParts> = {}
    for (const v of variants) { const p = splitHf((hf as Record<string, string | undefined>)[v.key]); o[v.key] = { left: hfToDisplay(p.left), center: hfToDisplay(p.center), right: hfToDisplay(p.right) } }
    return o
  })
  const refs = useRef<Record<HfSection, HTMLTextAreaElement | null>>({ left: null, center: null, right: null })
  const [focus, setFocus] = useState<HfSection>('center')
  const cur = variants[vi]
  const setPart = (s: HfSection, text: string): void => setParts((p) => ({ ...p, [cur.key]: { ...p[cur.key], [s]: text } }))
  const insert = (text: string): void => {
    const el = refs.current[focus]
    const v = parts[cur.key][focus]
    const a = el?.selectionStart ?? v.length, b = el?.selectionEnd ?? v.length
    setPart(focus, v.slice(0, a) + text + v.slice(b))
  }
  const wrap = (code: string): void => {
    const el = refs.current[focus]
    const v = parts[cur.key][focus]
    const a = el?.selectionStart ?? 0, b = el?.selectionEnd ?? v.length
    setPart(focus, v.slice(0, a) + code + v.slice(a, b) + code + v.slice(b))
  }
  const size = (pt: number): void => {
    const el = refs.current[focus]
    const v = parts[cur.key][focus]
    const a = el?.selectionStart ?? 0
    setPart(focus, v.slice(0, a) + `&${pt}` + v.slice(a))
  }
  const finish = (): void => {
    const out: NonNullable<PageSetup['hf']> = {}
    for (const v of variants) {
      const p = parts[v.key]
      ;(out as Record<string, string>)[v.key] = joinHf({ left: hfFromDisplay(p.left), center: hfFromDisplay(p.center), right: hfFromDisplay(p.right) })
    }
    onOk(out)
  }
  return (
    <Dialog title={`${kind}`} width={560} onOk={finish} onCancel={onCancel}>
      <div style={{ fontSize: 12 }}>
        {variants.length > 1 && (
          <select value={vi} onChange={(e) => setVi(+e.target.value)} style={{ marginBottom: 8 }}>
            {variants.map((v, i) => <option key={v.key} value={i}>{v.label}</option>)}
          </select>
        )}
        <div style={{ display: 'flex', gap: 6, marginBottom: 8, flexWrap: 'wrap' }}>
          <button className="xl-btn" title="Bold" style={{ fontWeight: 700 }} onClick={() => wrap('&B')}>B</button>
          <button className="xl-btn" title="Italic" style={{ fontStyle: 'italic' }} onClick={() => wrap('&I')}>I</button>
          <button className="xl-btn" title="Underline" style={{ textDecoration: 'underline' }} onClick={() => wrap('&U')}>U</button>
          <select title="Font size at cursor" value="" onChange={(e) => e.target.value && size(+e.target.value)}>
            <option value="">Size</option>{[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
          {INSERTS.map((i) => <button key={i.label} className="xl-btn" title={`Insert ${i.label}`} onClick={() => insert(i.text)}>{i.label}</button>)}
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          {(['left', 'center', 'right'] as HfSection[]).map((s) => (
            <div key={s} style={{ flex: 1 }}>
              <div>{s[0].toUpperCase() + s.slice(1)} section:</div>
              <textarea ref={(el) => { refs.current[s] = el }} rows={6} style={{ width: '100%', resize: 'none' }} value={parts[cur.key][s]} onFocus={() => setFocus(s)} onChange={(e) => setPart(s, e.target.value)} />
            </div>
          ))}
        </div>
      </div>
    </Dialog>
  )
}

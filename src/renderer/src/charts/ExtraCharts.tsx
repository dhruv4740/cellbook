// SVG bodies for the chart families beyond bar/line/pie/scatter: radar, funnel, stock, surface, treemap, sunburst,
// histogram, pareto and box & whisker. Styled like the other bodies in ChartRenderer.
import type { ReactElement } from 'react'
import type { ExtractedChartData } from './dataExtraction'
import { AXIS_COLOR, GRIDLINE_COLOR, TEXT_COLOR } from './theme'
import { boxStats, histogramBins, paretoOf, squarify, surfaceBands } from './extraStats'
import { axisScale, axisScaleRange } from './ChartRenderer'

type Box = { x: number; y: number; w: number; h: number }
interface Body { inner: Box; data: ExtractedChartData; colors: string[]; gridlines: boolean; axes: boolean; dataLabels: boolean }

const fmt = (v: number): string => String(Math.round(v * 1e4) / 1e4)
const trunc = (s: string, n = 12): string => (s.length > n ? s.slice(0, n - 1) + '…' : s)
const tx = (x: number, y: number, s: string, anchor: 'start' | 'middle' | 'end' = 'middle', size = 9.5, fill = TEXT_COLOR): ReactElement => (
  <text key={`${x}-${y}-${s}`} x={x} y={y} textAnchor={anchor} fontSize={size} fill={fill}>{s}</text>
)

function ValueTicks({ inner, min, max, ticks, gridlines, axes }: { inner: Box; min: number; max: number; ticks: number; gridlines: boolean; axes: boolean }) {
  return (
    <g>
      {Array.from({ length: ticks + 1 }, (_, i) => {
        const y = inner.y + inner.h - (inner.h * i) / ticks
        return (
          <g key={i}>
            {gridlines && <line x1={inner.x} x2={inner.x + inner.w} y1={y} y2={y} stroke={GRIDLINE_COLOR} />}
            {axes && tx(inner.x - 6, y + 3, fmt(min + ((max - min) * i) / ticks), 'end')}
          </g>
        )
      })}
      {axes && <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />}
    </g>
  )
}

export function RadarBody({ inner, data, colors, gridlines, axes, subtype }: Body & { subtype: string }) {
  const n = Math.max(3, data.categories.length)
  const cx = inner.x + inner.w / 2, cy = inner.y + inner.h / 2
  const r = Math.min(inner.w, inner.h) / 2 - 16
  let hi = 0
  data.series.forEach((s) => s.values.forEach((v) => { if (v !== null) hi = Math.max(hi, v) }))
  const { max, ticks } = axisScale(hi)
  const pt = (i: number, frac: number): [number, number] => [cx + r * frac * Math.sin((2 * Math.PI * i) / n), cy - r * frac * Math.cos((2 * Math.PI * i) / n)]
  const ring = (frac: number): string => Array.from({ length: n }, (_, i) => pt(i, frac).join(',')).join(' ')
  const filled = subtype === 'filledRadar', markers = subtype === 'radarMarkers'
  return (
    <g>
      {gridlines && Array.from({ length: ticks }, (_, k) => <polygon key={k} points={ring((k + 1) / ticks)} fill="none" stroke={GRIDLINE_COLOR} />)}
      {Array.from({ length: n }, (_, i) => <line key={i} x1={cx} y1={cy} x2={pt(i, 1)[0]} y2={pt(i, 1)[1]} stroke={GRIDLINE_COLOR} />)}
      {axes && Array.from({ length: ticks }, (_, k) => tx(cx - 4, cy - (r * (k + 1)) / ticks + 3, fmt((max * (k + 1)) / ticks), 'end', 8.5))}
      {data.series.map((s, si) => {
        const pts = Array.from({ length: n }, (_, i) => pt(i, Math.max(0, (s.values[i] ?? 0) / max)))
        const c = colors[si % colors.length]
        return (
          <g key={si}>
            <polygon points={pts.map((p) => p.join(',')).join(' ')} fill={filled ? c : 'none'} fillOpacity={filled ? 0.55 : 1} stroke={c} strokeWidth={2} />
            {markers && pts.map((p, i) => <circle key={i} cx={p[0]} cy={p[1]} r={3} fill={c} />)}
          </g>
        )
      })}
      {data.categories.slice(0, n).map((c, i) => {
        const [x, y] = pt(i, 1.1)
        return tx(x, y + 3, trunc(c), Math.abs(x - cx) < 4 ? 'middle' : x > cx ? 'start' : 'end')
      })}
    </g>
  )
}

export function FunnelBody({ inner, data, colors, dataLabels }: Body) {
  const vals = (data.series[0]?.values ?? []).map((v) => Math.max(0, v ?? 0))
  const n = Math.max(1, vals.length), max = Math.max(1, ...vals)
  const slot = inner.h / n
  const labelW = Math.min(90, inner.w * 0.25)
  const midX = inner.x + labelW + (inner.w - labelW) / 2, maxW = inner.w - labelW - 6
  return (
    <g>
      {vals.map((v, i) => {
        const w = (v / max) * maxW, y = inner.y + slot * i + slot * 0.08
        return (
          <g key={i}>
            <rect x={midX - w / 2} y={y} width={w} height={slot * 0.84} fill={colors[0]} />
            {tx(inner.x + labelW - 6, y + slot * 0.5 + 3, trunc(data.categories[i] ?? '', 14), 'end')}
            {dataLabels !== false && w > 24 && tx(midX, y + slot * 0.5 + 3, fmt(v), 'middle', 9.5, '#fff')}
          </g>
        )
      })}
    </g>
  )
}

/** Stock: series are High-Low-Close, or Open-High-Low-Close when there are four or more. */
export function StockBody({ inner, data, colors, gridlines, axes, subtype }: Body & { subtype: string }) {
  const ohlc = subtype === 'stockOHLC'
  const s = data.series
  const n = Math.max(1, data.categories.length)
  const get = (name: 'o' | 'h' | 'l' | 'c', i: number): number | null => {
    const idx = ohlc ? { o: 0, h: 1, l: 2, c: 3 }[name] : { o: -1, h: 0, l: 1, c: 2 }[name]
    return idx >= 0 ? (s[idx]?.values[i] ?? null) : null
  }
  let lo = Infinity, hi = -Infinity
  for (let i = 0; i < n; i++) for (const k of ['o', 'h', 'l', 'c'] as const) { const v = get(k, i); if (v !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v) } }
  if (!isFinite(lo)) { lo = 0; hi = 1 }
  const sc = axisScaleRange(lo > 0 ? 0 : lo, hi)
  const min = lo > 0 ? Math.floor(lo - (hi - lo) * 0.1) : sc.min
  const range = sc.max - min || 1
  const yOf = (v: number): number => inner.y + inner.h - ((v - min) / range) * inner.h
  const slot = inner.w / n
  return (
    <g>
      <ValueTicks inner={inner} min={min} max={sc.max} ticks={sc.ticks} gridlines={gridlines} axes={axes} />
      {Array.from({ length: n }, (_, i) => {
        const h = get('h', i), l = get('l', i), c = get('c', i), o = get('o', i)
        const x = inner.x + slot * (i + 0.5)
        return (
          <g key={i}>
            {h !== null && l !== null && <line x1={x} x2={x} y1={yOf(h)} y2={yOf(l)} stroke="#404040" strokeWidth={1.5} />}
            {ohlc && o !== null && c !== null && (
              <rect x={x - slot * 0.25} y={yOf(Math.max(o, c))} width={slot * 0.5} height={Math.max(1, Math.abs(yOf(o) - yOf(c)))} fill={c >= o ? '#fff' : '#404040'} stroke="#404040" />
            )}
            {!ohlc && c !== null && <line x1={x} x2={x + slot * 0.25} y1={yOf(c)} y2={yOf(c)} stroke="#404040" strokeWidth={1.5} />}
          </g>
        )
      })}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          {data.categories.map((c, i) => (n > 12 && i % Math.ceil(n / 12) ? null : tx(inner.x + slot * (i + 0.5), inner.y + inner.h + 13, trunc(c, 8))))}
        </>
      )}
      <g fill={colors[0]} />
    </g>
  )
}

const BAND_COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

/** Surface: series are rows, categories columns. 3-D surface is an isometric projection; contour views look straight down. */
export function SurfaceBody({ inner, data, subtype }: Body & { subtype: string }) {
  const rows = data.series.length, cols = Math.max(1, data.categories.length)
  let lo = Infinity, hi = -Infinity
  data.series.forEach((s) => s.values.forEach((v) => { if (v !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v) } }))
  if (!isFinite(lo)) { lo = 0; hi = 1 }
  const bands = surfaceBands(lo, hi)
  const bandOf = (v: number): number => Math.min(bands.length - 1, Math.max(0, Math.floor(((v - lo) / ((hi - lo) || 1)) * bands.length)))
  const val = (r: number, c: number): number => data.series[r]?.values[c] ?? lo
  const wire = subtype === 'wireSurface' || subtype === 'wireContour'
  const top = subtype === 'contour' || subtype === 'wireContour'
  const legendW = 70
  const area: Box = { x: inner.x, y: inner.y, w: inner.w - legendW, h: inner.h }
  const quads: ReactElement[] = []
  if (top) {
    const cw = area.w / cols, ch = area.h / Math.max(1, rows)
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const b = bandOf(val(r, c))
      quads.push(<rect key={`${r}-${c}`} x={area.x + c * cw} y={area.y + r * ch} width={cw} height={ch} fill={wire ? 'none' : BAND_COLORS[b]} stroke={wire ? BAND_COLORS[b] : '#fff'} strokeWidth={wire ? 1.5 : 0.5} />)
    }
  } else {
    const dx = area.w / (cols + rows * 0.5), dy = area.h / 3
    const proj = (r: number, c: number, v: number): [number, number] => [
      area.x + c * dx + (rows - 1 - r) * dx * 0.5 + 6,
      area.y + area.h - (rows - 1 - r) * dy * 0.25 - 4 - ((v - lo) / ((hi - lo) || 1)) * dy * 2
    ]
    for (let r = 0; r < rows - 1; r++) for (let c = 0; c < cols - 1; c++) { // back row first so nearer faces paint over it
      const pts = [proj(r, c, val(r, c)), proj(r, c + 1, val(r, c + 1)), proj(r + 1, c + 1, val(r + 1, c + 1)), proj(r + 1, c, val(r + 1, c))]
      const b = bandOf((val(r, c) + val(r, c + 1) + val(r + 1, c) + val(r + 1, c + 1)) / 4)
      quads.push(<polygon key={`${r}-${c}`} points={pts.map((p) => p.join(',')).join(' ')} fill={wire ? 'none' : BAND_COLORS[b]} stroke={wire ? '#404040' : '#fff'} strokeWidth={wire ? 0.8 : 0.5} />)
    }
    if (rows < 2 || cols < 2) quads.push(tx(area.x + area.w / 2, area.y + area.h / 2, 'A surface needs at least 2 series and 2 categories.'))
  }
  return (
    <g>
      {quads}
      {bands.map((b, i) => (
        <g key={i} transform={`translate(${inner.x + inner.w - legendW + 6}, ${inner.y + inner.h / 2 - bands.length * 8 + i * 16})`}>
          <rect width={9} height={9} y={1} fill={BAND_COLORS[i]} />
          {tx(13, 9, `${fmt(b.lo)}-${fmt(b.hi)}`, 'start', 9)}
        </g>
      ))}
    </g>
  )
}

export function TreemapBody({ inner, data, colors, dataLabels }: Body) {
  const vals = (data.series[0]?.values ?? []).map((v) => v ?? 0)
  const tiles = squarify(vals, inner)
  return (
    <g>
      {tiles.map((t) => (
        <g key={t.i}>
          <rect x={t.x} y={t.y} width={t.w} height={t.h} fill={colors[t.i % colors.length]} stroke="#fff" strokeWidth={2} />
          {t.w > 34 && t.h > 18 && tx(t.x + 5, t.y + 13, trunc(data.categories[t.i] ?? '', Math.max(4, Math.floor(t.w / 6))), 'start', 10, '#fff')}
          {dataLabels && t.w > 34 && t.h > 32 && tx(t.x + 5, t.y + 26, fmt(vals[t.i]), 'start', 9, '#fff')}
        </g>
      ))}
    </g>
  )
}

/** Sunburst: one ring per series (first series innermost); slices follow each series' values and keep their category colour. */
export function SunburstBody({ inner, data, colors }: Body) {
  const cx = inner.x + inner.w / 2, cy = inner.y + inner.h / 2
  const R = Math.min(inner.w, inner.h) / 2 - 4
  const rings = Math.max(1, data.series.length)
  const arcs: ReactElement[] = []
  data.series.forEach((s, ring) => {
    const vals = s.values.map((v) => Math.max(0, v ?? 0))
    const total = vals.reduce((a, b) => a + b, 0) || 1
    const r0 = (R * ring) / rings + (ring === 0 ? 0 : 0), r1 = (R * (ring + 1)) / rings
    let a = -Math.PI / 2
    vals.forEach((v, i) => {
      const a1 = a + (v / total) * Math.PI * 2
      const p = (ang: number, rad: number): string => `${cx + rad * Math.cos(ang)} ${cy + rad * Math.sin(ang)}`
      const large = a1 - a > Math.PI ? 1 : 0
      if (v > 0) {
        const full = a1 - a > Math.PI * 2 - 1e-6
        const d = full
          ? `M${cx - r1} ${cy} A${r1} ${r1} 0 1 1 ${cx + r1} ${cy} A${r1} ${r1} 0 1 1 ${cx - r1} ${cy} ${r0 > 0 ? `M${cx - r0} ${cy} A${r0} ${r0} 0 1 0 ${cx + r0} ${cy} A${r0} ${r0} 0 1 0 ${cx - r0} ${cy}` : ''}`
          : `M${p(a, r0)} L${p(a, r1)} A${r1} ${r1} 0 ${large} 1 ${p(a1, r1)} L${p(a1, r0)} A${r0} ${r0} 0 ${large} 0 ${p(a, r0)} Z`
        arcs.push(<path key={`${ring}-${i}`} d={d} fill={colors[i % colors.length]} stroke="#fff" strokeWidth={1.5} fillRule="evenodd" />)
        const mid = (a + a1) / 2, rm = (r0 + r1) / 2
        if (a1 - a > 0.25 && ring === 0) arcs.push(tx(cx + rm * Math.cos(mid), cy + rm * Math.sin(mid) + 3, trunc(data.categories[i] ?? '', 8), 'middle', 9, '#fff'))
      }
      a = a1
    })
  })
  return <g>{arcs}</g>
}

function CountAxes({ inner, labels, max, ticks, gridlines, axes }: { inner: Box; labels: string[]; max: number; ticks: number; gridlines: boolean; axes: boolean }) {
  const n = Math.max(1, labels.length), slot = inner.w / n, skip = n > 10 ? Math.ceil(n / 10) : 1
  return (
    <g>
      <ValueTicks inner={inner} min={0} max={max} ticks={ticks} gridlines={gridlines} axes={axes} />
      {axes && <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />}
      {axes && labels.map((l, i) => (i % skip ? null : tx(inner.x + slot * (i + 0.5), inner.y + inner.h + 13, l, 'middle', 8.5)))}
    </g>
  )
}

export function HistogramBody({ inner, data, colors, gridlines, axes, dataLabels }: Body) {
  const bins = histogramBins(data.series[0]?.values ?? [])
  const { max, ticks } = axisScale(Math.max(1, ...bins.map((b) => b.count)))
  const slot = inner.w / Math.max(1, bins.length)
  return (
    <g>
      <CountAxes inner={inner} labels={bins.map((b) => b.label)} max={max} ticks={ticks} gridlines={gridlines} axes={axes} />
      {bins.map((b, i) => {
        const h = (b.count / max) * inner.h
        return (
          <g key={i}>
            <rect x={inner.x + slot * i + slot * 0.025} y={inner.y + inner.h - h} width={slot * 0.95} height={h} fill={colors[0]} />
            {dataLabels && b.count > 0 && tx(inner.x + slot * (i + 0.5), inner.y + inner.h - h - 3, String(b.count), 'middle', 9)}
          </g>
        )
      })}
    </g>
  )
}

export function ParetoBody({ inner, data, colors, gridlines, axes, dataLabels }: Body) {
  const rows = paretoOf(histogramBins(data.series[0]?.values ?? []))
  const { max, ticks } = axisScale(Math.max(1, ...rows.map((b) => b.count)))
  const slot = inner.w / Math.max(1, rows.length)
  const plot = { ...inner, w: inner.w - 30 }
  const pslot = plot.w / Math.max(1, rows.length)
  const line = rows.map((r, i) => `${plot.x + pslot * (i + 0.5)},${plot.y + plot.h - (r.cumPct / 100) * plot.h}`).join(' ')
  void slot
  return (
    <g>
      <CountAxes inner={plot} labels={rows.map((r) => r.label)} max={max} ticks={ticks} gridlines={gridlines} axes={axes} />
      {rows.map((r, i) => {
        const h = (r.count / max) * plot.h
        return (
          <g key={i}>
            <rect x={plot.x + pslot * i + pslot * 0.025} y={plot.y + plot.h - h} width={pslot * 0.95} height={h} fill={colors[0]} />
            {dataLabels && r.count > 0 && tx(plot.x + pslot * (i + 0.5), plot.y + plot.h - h - 3, String(r.count), 'middle', 9)}
          </g>
        )
      })}
      <polyline points={line} fill="none" stroke={colors[1 % colors.length]} strokeWidth={2} />
      {rows.map((r, i) => <circle key={i} cx={plot.x + pslot * (i + 0.5)} cy={plot.y + plot.h - (r.cumPct / 100) * plot.h} r={3} fill={colors[1 % colors.length]} />)}
      {axes && [0, 25, 50, 75, 100].map((p) => tx(plot.x + plot.w + 4, plot.y + plot.h - (p / 100) * plot.h + 3, p + '%', 'start', 9))}
    </g>
  )
}

export function BoxBody({ inner, data, colors, gridlines, axes }: Body) {
  const stats = data.series.map((s) => boxStats(s.values))
  let lo = Infinity, hi = -Infinity
  stats.forEach((b) => { if (b) { lo = Math.min(lo, b.whiskerLo, ...b.outliers); hi = Math.max(hi, b.whiskerHi, ...b.outliers) } })
  if (!isFinite(lo)) { lo = 0; hi = 1 }
  const sc = axisScaleRange(Math.min(0, lo), hi)
  const range = sc.max - sc.min || 1
  const yOf = (v: number): number => inner.y + inner.h - ((v - sc.min) / range) * inner.h
  const slot = inner.w / Math.max(1, stats.length)
  return (
    <g>
      <ValueTicks inner={inner} min={sc.min} max={sc.max} ticks={sc.ticks} gridlines={gridlines} axes={axes} />
      {axes && <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />}
      {stats.map((b, i) => {
        if (!b) return null
        const x = inner.x + slot * (i + 0.5), w = slot * 0.5, c = colors[i % colors.length]
        return (
          <g key={i}>
            <line x1={x} x2={x} y1={yOf(b.whiskerHi)} y2={yOf(b.q3)} stroke={c} strokeWidth={1.5} />
            <line x1={x} x2={x} y1={yOf(b.q1)} y2={yOf(b.whiskerLo)} stroke={c} strokeWidth={1.5} />
            <line x1={x - w / 4} x2={x + w / 4} y1={yOf(b.whiskerHi)} y2={yOf(b.whiskerHi)} stroke={c} strokeWidth={1.5} />
            <line x1={x - w / 4} x2={x + w / 4} y1={yOf(b.whiskerLo)} y2={yOf(b.whiskerLo)} stroke={c} strokeWidth={1.5} />
            <rect x={x - w / 2} y={yOf(b.q3)} width={w} height={Math.max(1, yOf(b.q1) - yOf(b.q3))} fill={c} fillOpacity={0.5} stroke={c} strokeWidth={1.5} />
            <line x1={x - w / 2} x2={x + w / 2} y1={yOf(b.median)} y2={yOf(b.median)} stroke={c} strokeWidth={2} />
            <path d={`M${x - 4} ${yOf(b.mean) - 4} L${x + 4} ${yOf(b.mean) + 4} M${x + 4} ${yOf(b.mean) - 4} L${x - 4} ${yOf(b.mean) + 4}`} stroke="#fff" strokeWidth={1.5} />
            {b.outliers.map((o, k) => <circle key={k} cx={x} cy={yOf(o)} r={3} fill="none" stroke={c} />)}
          </g>
        )
      })}
    </g>
  )
}

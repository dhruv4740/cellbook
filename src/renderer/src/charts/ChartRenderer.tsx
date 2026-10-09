// Draws a ChartSpec + its extracted data as inline SVG, styled like an Office 2023+ default chart
// (Calibri font, theme accent colors, light gridlines, "Chart Title" placeholder).
import type { ReactElement } from 'react'
import { ChartSpec } from '../model/types'
import { ExtractedChartData } from './dataExtraction'
import { CHART_FONT, GRIDLINE_COLOR, AXIS_COLOR, TITLE_COLOR, TEXT_COLOR, colorSet } from './theme'
import { comboSeriesOf, niceRange, waterfallSteps } from './chartCombo'
import { RadarBody, FunnelBody, StockBody, SurfaceBody, TreemapBody, SunburstBody, HistogramBody, ParetoBody, BoxBody } from './ExtraCharts'

export interface ChartRendererProps {
  spec: ChartSpec
  data: ExtractedChartData
  width: number
  height: number
}

const isBarLike = (t: string) => t === 'bar'
const isStacked = (sub: string) => sub.startsWith('stacked') || sub.startsWith('percentStacked')
const isPercent = (sub: string) => sub.startsWith('percentStacked')
const isHorizontal = (sub: string) => sub === 'clusteredBar' || sub === 'stackedBar' || sub === 'percentStackedBar'
const isArea = (sub: string) => sub === 'area' || sub === 'stackedArea'
const isLineMarkers = (sub: string) => sub === 'lineMarkers'
const isDoughnut = (sub: string) => sub === 'doughnut'
const isSmooth = (sub: string) => sub === 'scatterSmooth'

/** Excel-style automatic value axis: a 1/2/5 x 10^k major unit giving about 5-10 intervals,
 *  and one extra unit of headroom when the data max lands on or near the top gridline
 *  (e.g. data max 5 -> axis 0..6 step 1, as Excel draws it). */
export function axisScale(v: number): { max: number; ticks: number } {
  if (!(v > 0)) return { max: 1, ticks: 5 }
  const raw = v / 5
  const mag = Math.pow(10, Math.floor(Math.log10(raw)))
  const norm = raw / mag
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag
  let ticks = Math.ceil(v / step - 1e-9)
  if (ticks * step - v < step * 0.05) ticks++
  return { max: ticks * step, ticks }
}

/** Excel-style automatic axis for data spanning [min, max]. Non-negative data gives exactly `axisScale` (axis from 0);
 *  with negatives the axis runs from a negative multiple of a 1/2/5 x 10^k step to a positive one (about 8 intervals at most),
 *  with one extra unit when the data lands on or near an end (e.g. -30..100 -> -40..120 step 20). */
export function axisScaleRange(min: number, max: number): { min: number; max: number; ticks: number; step: number } {
  if (!(min < 0)) {
    const s = axisScale(max)
    return { min: 0, max: s.max, ticks: s.ticks, step: s.max / s.ticks }
  }
  const hi = Math.max(0, max)
  const span = hi - min
  for (let k = Math.floor(Math.log10(span / 50)); ; k++) {
    for (const m of [1, 2, 5]) {
      const step = m * Math.pow(10, k)
      let up = hi > 0 ? Math.ceil(hi / step - 1e-9) : 0
      if (hi > 0 && up * step - hi < step * 0.05) up++
      let down = Math.ceil(-min / step - 1e-9)
      if (down * step + min < step * 0.05) down++
      if (up + down <= 8) return { min: -down * step, max: up * step, ticks: up + down, step }
    }
  }
}

export function ChartRenderer({ spec, data, width, height }: ChartRendererProps) {
  // Pie/doughnut charts have one series but color each category slice separately.
  const byPoint = spec.type === 'pie' || spec.type === 'funnel' || spec.type === 'hierarchy'
  const colorCount = byPoint ? data.categories.length : data.series.length
  const colors = colorSet(spec.colorSetId ?? 0, Math.max(1, colorCount || 1))
  const el = spec.elements
  const titleH = el.title ? 26 : 6
  const legendPos = el.legend ? spec.legendPosition : 'none'
  let legendW = 0, legendH = 0
  if (legendPos === 'right' || legendPos === 'left') legendW = Math.min(120, Math.max(70, width * 0.28))
  if (legendPos === 'top' || legendPos === 'bottom') legendH = 22

  const plot = {
    x: 12 + (legendPos === 'left' ? legendW : 0),
    y: titleH + (legendPos === 'top' ? legendH : 0) + 8,
    w: width - 24 - legendW,
    h: height - titleH - legendH - 18
  }

  const isPie = spec.type === 'pie'
  const isScatter = spec.type === 'scatter'
  const axisPad = { left: isPie ? 0 : 34, bottom: isPie || isScatter === false ? 18 : 18 }
  const inner = isPie
    ? plot
    : { x: plot.x + axisPad.left, y: plot.y, w: plot.w - axisPad.left, h: plot.h - axisPad.bottom }

  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ fontFamily: CHART_FONT, display: 'block' }}>
      <rect x={0} y={0} width={width} height={height} fill="#FFFFFF" />
      {el.title && (
        <text x={width / 2} y={18} textAnchor="middle" fontSize={13} fill={TITLE_COLOR}>
          {spec.title || 'Chart Title'}
        </text>
      )}
      {isPie ? (
        <PieBody inner={inner} data={data} colors={colors} doughnut={isDoughnut(spec.subtype)} dataLabels={el.dataLabels} />
      ) : isScatter ? (
        <ScatterBody inner={inner} data={data} colors={colors} smooth={isSmooth(spec.subtype)} gridlines={el.gridlines} axes={el.axes} />
      ) : spec.type === 'combo' ? (
        <ComboBody inner={inner} data={data} colors={colors} spec={spec} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.type === 'radar' ? (
        <RadarBody inner={inner} data={data} colors={colors} subtype={spec.subtype} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.type === 'funnel' ? (
        <FunnelBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.type === 'stock' ? (
        <StockBody inner={inner} data={data} colors={colors} subtype={spec.subtype} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.type === 'surface' ? (
        <SurfaceBody inner={inner} data={data} colors={colors} subtype={spec.subtype} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.subtype === 'treemap' ? (
        <TreemapBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.subtype === 'sunburst' ? (
        <SunburstBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.subtype === 'histogram' ? (
        <HistogramBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.subtype === 'pareto' ? (
        <ParetoBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.subtype === 'boxWhisker' ? (
        <BoxBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.type === 'waterfall' ? (
        <WaterfallBody inner={inner} data={data} colors={colors} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : spec.type === 'line' ? (
        <LineBody inner={inner} data={data} colors={colors} subtype={spec.subtype} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      ) : (
        <BarBody inner={inner} data={data} colors={colors} subtype={spec.subtype} gridlines={el.gridlines} axes={el.axes} dataLabels={el.dataLabels} />
      )}
      {el.legend && legendPos !== 'none' && (
        <Legend pos={legendPos} width={width} height={height} legendW={legendW} legendH={legendH} plot={plot} names={legendNames(spec, data)} colors={colors} />
      )}
    </svg>
  )
}

function legendNames(spec: ChartSpec, data: ExtractedChartData): string[] {
  if (spec.type === 'pie') return data.categories
  if (spec.type === 'waterfall') return ['Increase', 'Decrease']
  if (spec.type === 'funnel' || spec.type === 'hierarchy') return data.categories
  return data.series.map((s) => s.name)
}

function Legend({ pos, width, height, legendW, legendH, plot, names, colors }: {
  pos: 'right' | 'left' | 'top' | 'bottom'
  width: number; height: number; legendW: number; legendH: number
  plot: { x: number; y: number; w: number; h: number }
  names: string[]; colors: string[]
}) {
  const items = names.slice(0, 12)
  if (pos === 'right' || pos === 'left') {
    const x0 = pos === 'right' ? width - legendW + 6 : 6
    const y0 = height / 2 - (items.length * 16) / 2
    return (
      <g fontSize={10.5} fill={TEXT_COLOR}>
        {items.map((n, i) => (
          <g key={i} transform={`translate(${x0}, ${y0 + i * 16})`}>
            <rect width={9} height={9} y={1} fill={colors[i % colors.length]} />
            <text x={13} y={9}>{n}</text>
          </g>
        ))}
      </g>
    )
  }
  const y0 = pos === 'top' ? 20 : height - 14
  const itemW = Math.min(110, plot.w / Math.max(1, items.length))
  const totalW = items.length * itemW
  const x0 = (width - totalW) / 2
  return (
    <g fontSize={10.5} fill={TEXT_COLOR}>
      {items.map((n, i) => (
        <g key={i} transform={`translate(${x0 + i * itemW}, ${y0})`}>
          <rect width={9} height={9} y={-8} fill={colors[i % colors.length]} />
          <text x={13} y={0}>{n.length > 14 ? n.slice(0, 13) + '…' : n}</text>
        </g>
      ))}
    </g>
  )
}

function Gridlines({ inner, ticks }: { inner: { x: number; y: number; w: number; h: number }; ticks: number }) {
  const lines = []
  for (let i = 0; i <= ticks; i++) {
    const y = inner.y + (inner.h * i) / ticks
    lines.push(<line key={i} x1={inner.x} x2={inner.x + inner.w} y1={y} y2={y} stroke={GRIDLINE_COLOR} strokeWidth={1} />)
  }
  return <g>{lines}</g>
}

function ValueAxis({ inner, min = 0, max, ticks, percent }: { inner: { x: number; y: number; w: number; h: number }; min?: number; max: number; ticks: number; percent?: boolean }) {
  const labels = []
  for (let i = 0; i <= ticks; i++) {
    const v = min + ((max - min) * i) / ticks
    const y = inner.y + inner.h - (inner.h * i) / ticks
    labels.push(
      <text key={i} x={inner.x - 6} y={y + 3} textAnchor="end" fontSize={9.5} fill={TEXT_COLOR}>
        {percent ? Math.round(v) + '%' : formatAxisNum(v)}
      </text>
    )
  }
  return <g>{labels}</g>
}

function formatAxisNum(v: number): string {
  // Excel's General axis format with the source column's thousands separator.
  return (Math.round(v * 1e6) / 1e6).toLocaleString('en-US', { maximumFractionDigits: 6 })
}

function CategoryAxis({ inner, categories, horizontal }: { inner: { x: number; y: number; w: number; h: number }; categories: string[]; horizontal?: boolean }) {
  const n = Math.max(1, categories.length)
  if (horizontal) {
    return (
      <g>
        {categories.map((c, i) => (
          <text key={i} x={inner.x - 6} y={inner.y + (inner.h * (i + 0.5)) / n + 3} textAnchor="end" fontSize={9.5} fill={TEXT_COLOR}>
            {trim(c)}
          </text>
        ))}
      </g>
    )
  }
  const step = inner.w / n
  const skip = categories.length > 12 ? Math.ceil(categories.length / 12) : 1
  return (
    <g>
      {categories.map((c, i) =>
        i % skip === 0 ? (
          <text key={i} x={inner.x + step * (i + 0.5)} y={inner.y + inner.h + 13} textAnchor="middle" fontSize={9.5} fill={TEXT_COLOR}>
            {trim(c)}
          </text>
        ) : null
      )}
    </g>
  )
}

function trim(s: string): string {
  return s.length > 10 ? s.slice(0, 9) + '…' : s
}

function BarBody({ inner, data, colors, subtype, gridlines, axes, dataLabels }: {
  inner: { x: number; y: number; w: number; h: number }
  data: ExtractedChartData; colors: string[]; subtype: string; gridlines: boolean; axes: boolean; dataLabels: boolean
}) {
  const stacked = isStacked(subtype)
  const percent = isPercent(subtype)
  const horizontal = isHorizontal(subtype)
  const n = Math.max(1, data.categories.length)
  const sCount = Math.max(1, data.series.length)

  // percent charts divide by the category's absolute total; negatives stack downward separately from positives
  const valAt = (si: number, i: number): number => {
    const raw = data.series[si].values[i] ?? 0
    if (!percent) return raw
    const total = data.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0)
    return total > 0 ? (raw / total) * 100 : raw
  }
  let lo = 0, hi = 0
  for (let i = 0; i < n; i++) {
    let pos = 0, neg = 0
    for (let si = 0; si < data.series.length; si++) {
      const v = valAt(si, i)
      if (stacked) { if (v < 0) neg += v; else pos += v } else { pos = Math.max(pos, v); neg = Math.min(neg, v) }
    }
    hi = Math.max(hi, pos); lo = Math.min(lo, neg)
  }
  if (lo >= 0) hi = Math.max(1, hi)
  const scale = percent && lo >= 0 ? { min: 0, max: 100, ticks: 10 } : axisScaleRange(lo, hi)
  const { min, max } = scale
  const toPx = (v: number): number => ((v - min) / (max - min)) * (horizontal ? inner.w : inner.h)

  const bars: ReactElement[] = []
  const catStep = (horizontal ? inner.h : inner.w) / n
  const groupPad = catStep * 0.18
  const groupW = catStep - groupPad * 2
  const barW = stacked ? groupW : groupW / sCount

  for (let i = 0; i < n; i++) {
    let posAcc = 0, negAcc = 0
    for (let si = 0; si < data.series.length; si++) {
      const raw = data.series[si].values[i] ?? 0
      const val = valAt(si, i)
      const from = stacked ? (val < 0 ? negAcc : posAcc) : 0
      const to = from + val
      if (stacked) { if (val < 0) negAcc = to; else posAcc = to }
      const a = toPx(Math.min(from, to)), b = toPx(Math.max(from, to))
      const len = b - a
      if (!horizontal) {
        const x = inner.x + catStep * i + groupPad + (stacked ? 0 : barW * si)
        const y = inner.y + inner.h - b
        bars.push(<rect key={si + '-' + i} x={x} y={y} width={Math.max(0, barW - 2)} height={Math.max(0, len)} fill={colors[si % colors.length]} />)
        if (dataLabels && len > 10) bars.push(<text key={'l' + si + '-' + i} x={x + barW / 2 - 1} y={val < 0 ? y + len + 10 : y - 3} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(raw)}</text>)
      } else {
        const y = inner.y + catStep * i + groupPad + (stacked ? 0 : barW * si)
        const x = inner.x + a
        bars.push(<rect key={si + '-' + i} x={x} y={y} width={Math.max(0, len)} height={Math.max(0, barW - 2)} fill={colors[si % colors.length]} />)
      }
    }
  }
  const zero = toPx(0)

  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={scale.ticks} />}
      {bars}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          {min < 0 && (horizontal
            ? <line x1={inner.x + zero} x2={inner.x + zero} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
            : <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h - zero} y2={inner.y + inner.h - zero} stroke={AXIS_COLOR} />)}
          {!horizontal && <ValueAxis inner={inner} min={min} max={max} ticks={scale.ticks} percent={percent} />}
          {horizontal && <CategoryAxis inner={inner} categories={data.categories} horizontal />}
          {!horizontal && <CategoryAxis inner={inner} categories={data.categories} />}
        </>
      )}
    </g>
  )
}

function fmtVal(v: number): string {
  return String(Math.round(v * 1e6) / 1e6)
}

function pathFor(points: [number, number][], smooth: boolean): string {
  if (!points.length) return ''
  if (!smooth || points.length < 3) return points.map((p, i) => (i === 0 ? 'M' : 'L') + p[0] + ' ' + p[1]).join(' ')
  let d = 'M' + points[0][0] + ' ' + points[0][1]
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[i], p1 = points[i + 1]
    const mx = (p0[0] + p1[0]) / 2
    d += ` Q${p0[0]} ${p0[1]} ${mx} ${(p0[1] + p1[1]) / 2}`
  }
  d += ` L${points[points.length - 1][0]} ${points[points.length - 1][1]}`
  return d
}

function LineBody({ inner, data, colors, subtype, gridlines, axes, dataLabels }: {
  inner: { x: number; y: number; w: number; h: number }
  data: ExtractedChartData; colors: string[]; subtype: string; gridlines: boolean; axes: boolean; dataLabels: boolean
}) {
  const stacked = isStacked(subtype)
  const area = isArea(subtype)
  const markers = isLineMarkers(subtype)
  const n = Math.max(1, data.categories.length)
  let lo = 0, hi = 0
  const run = new Array(n).fill(0)
  for (const s of data.series) {
    s.values.forEach((v, i) => {
      const t = stacked ? (run[i] += v ?? 0) : (v ?? 0)
      lo = Math.min(lo, t); hi = Math.max(hi, t)
    })
  }
  if (lo >= 0) hi = Math.max(1, hi)
  const scale = axisScaleRange(lo, hi)
  const { min, max } = scale
  const yOf = (v: number): number => inner.y + inner.h - ((v - min) / (max - min)) * inner.h
  const step = n > 1 ? inner.w / (n - 1) : inner.w

  const stackAcc = new Array(n).fill(0)
  const shapes: ReactElement[] = []
  data.series.forEach((s, si) => {
    const pts: [number, number][] = s.values.map((v, i) => {
      const raw = v ?? 0
      if (stacked) stackAcc[i] += raw
      return [inner.x + step * i, yOf(stacked ? stackAcc[i] : raw)]
    })
    const d = pathFor(pts, false)
    if (area) {
      const baseY = yOf(0)
      const areaD = `${d} L${pts[pts.length - 1][0]} ${baseY} L${pts[0][0]} ${baseY} Z`
      shapes.push(<path key={'a' + si} d={areaD} fill={colors[si % colors.length]} opacity={0.55} stroke="none" />)
    }
    // empty cells leave a gap in a plain line (Excel's default), rather than dropping to zero
    const runs: [number, number][][] = [[]]
    pts.forEach((pt, i) => { if (s.values[i] === null && !stacked && !area) runs.push([]); else runs[runs.length - 1].push(pt) })
    runs.forEach((run, ri) => { if (run.length > 1) shapes.push(<path key={si + '_' + ri} d={pathFor(run, false)} fill="none" stroke={colors[si % colors.length]} strokeWidth={2} />) })
    if (markers) pts.forEach(([x, y], i) => s.values[i] !== null && shapes.push(<circle key={si + '-' + i} cx={x} cy={y} r={3} fill="#FFFFFF" stroke={colors[si % colors.length]} strokeWidth={1.5} />))
    if (dataLabels) pts.forEach(([x, y], i) => shapes.push(<text key={'l' + si + '-' + i} x={x} y={y - 6} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(s.values[i] ?? 0)}</text>))
  })

  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={scale.ticks} />}
      {shapes}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          {min < 0 && <line x1={inner.x} x2={inner.x + inner.w} y1={yOf(0)} y2={yOf(0)} stroke={AXIS_COLOR} />}
          <ValueAxis inner={inner} min={min} max={max} ticks={scale.ticks} />
          <CategoryAxis inner={inner} categories={data.categories} />
        </>
      )}
    </g>
  )
}

function ScatterBody({ inner, data, colors, smooth, gridlines, axes }: {
  inner: { x: number; y: number; w: number; h: number }
  data: ExtractedChartData; colors: string[]; smooth: boolean; gridlines: boolean; axes: boolean
}) {
  // Approximation: category index serves as X (a numeric X series isn't distinguished from Y series
  // by the generic extractor), each chart series plots as Y.
  const n = Math.max(1, data.categories.length)
  let maxY = 1
  for (const s of data.series) for (const v of s.values) maxY = Math.max(maxY, v ?? 0)
  const scale = axisScale(maxY)
  maxY = scale.max
  const step = n > 1 ? inner.w / (n - 1) : inner.w

  const shapes: ReactElement[] = []
  data.series.forEach((s, si) => {
    const pts: [number, number][] = s.values.map((v, i) => [inner.x + step * i, inner.y + inner.h - (Math.max(0, v ?? 0) / maxY) * inner.h])
    if (smooth) shapes.push(<path key={'p' + si} d={pathFor(pts, true)} fill="none" stroke={colors[si % colors.length]} strokeWidth={2} />)
    pts.forEach(([x, y], i) => shapes.push(<circle key={si + '-' + i} cx={x} cy={y} r={3.5} fill={colors[si % colors.length]} />))
  })

  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={scale.ticks} />}
      {shapes}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <ValueAxis inner={inner} max={maxY} ticks={scale.ticks} />
          <CategoryAxis inner={inner} categories={data.categories} />
        </>
      )}
    </g>
  )
}

function ComboBody({ inner: box, data, colors, spec, gridlines, axes, dataLabels }: {
  inner: { x: number; y: number; w: number; h: number }
  data: ExtractedChartData; colors: string[]; spec: ChartSpec; gridlines: boolean; axes: boolean; dataLabels: boolean
}) {
  const kinds = comboSeriesOf(spec, data.series.length)
  const hasSecondary = kinds.some((k, i) => k.secondary && i < data.series.length)
  const inner = hasSecondary ? { ...box, w: box.w - 30 } : box
  const n = Math.max(1, data.categories.length)
  const rangeOf = (secondary: boolean): ReturnType<typeof axisScaleRange> => {
    let lo = 0, hi = 0
    data.series.forEach((s, i) => { if (!!kinds[i].secondary === secondary) for (const v of s.values) { lo = Math.min(lo, v ?? 0); hi = Math.max(hi, v ?? 0) } })
    if (lo >= 0) hi = Math.max(1, hi)
    return axisScaleRange(lo, hi)
  }
  const scales = [rangeOf(false), rangeOf(true)]
  const step = inner.w / n
  const barSeries = data.series.map((_, i) => i).filter((i) => kinds[i].type === 'bar')
  const groupPad = step * 0.18
  const barW = (step - groupPad * 2) / Math.max(1, barSeries.length)
  const shapes: ReactElement[] = []
  const lineShapes: ReactElement[] = []
  data.series.forEach((s, si) => {
    const sc = scales[kinds[si].secondary ? 1 : 0]
    const yOf = (v: number): number => inner.y + inner.h - ((v - sc.min) / (sc.max - sc.min)) * inner.h
    const color = colors[si % colors.length]
    if (kinds[si].type === 'bar') {
      const slot = barSeries.indexOf(si)
      s.values.forEach((v, i) => {
        const y = yOf(v ?? 0), y0 = yOf(0)
        const x = inner.x + step * i + groupPad + barW * slot
        shapes.push(<rect key={si + '-' + i} x={x} y={Math.min(y, y0)} width={Math.max(0, barW - 2)} height={Math.abs(y0 - y)} fill={color} />)
        if (dataLabels && v !== null) shapes.push(<text key={'l' + si + '-' + i} x={x + barW / 2 - 1} y={v < 0 ? y + 10 : y - 3} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(v)}</text>)
      })
    } else {
      const pts: [number, number][] = s.values.map((v, i) => [inner.x + step * (i + 0.5), yOf(v ?? 0)])
      const runs: [number, number][][] = [[]]
      pts.forEach((pt, i) => { if (s.values[i] === null) runs.push([]); else runs[runs.length - 1].push(pt) })
      runs.forEach((run, ri) => { if (run.length > 1) lineShapes.push(<path key={si + '_' + ri} d={pathFor(run, false)} fill="none" stroke={color} strokeWidth={2} />) })
      pts.forEach(([x, y], i) => {
        if (s.values[i] === null) return
        lineShapes.push(<circle key={si + '-' + i} cx={x} cy={y} r={3} fill="#FFFFFF" stroke={color} strokeWidth={1.5} />)
        if (dataLabels) lineShapes.push(<text key={'l' + si + '-' + i} x={x} y={y - 6} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(s.values[i] ?? 0)}</text>)
      })
    }
  })
  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={scales[0].ticks} />}
      {shapes}
      {lineShapes}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          {scales[0].min < 0 && <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h - ((0 - scales[0].min) / (scales[0].max - scales[0].min)) * inner.h} y2={inner.y + inner.h - ((0 - scales[0].min) / (scales[0].max - scales[0].min)) * inner.h} stroke={AXIS_COLOR} />}
          <ValueAxis inner={inner} min={scales[0].min} max={scales[0].max} ticks={scales[0].ticks} />
          <CategoryAxis inner={inner} categories={data.categories} />
          {hasSecondary && (
            <g>
              <line x1={inner.x + inner.w} x2={inner.x + inner.w} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
              {Array.from({ length: scales[1].ticks + 1 }, (_, i) => (
                <text key={i} x={inner.x + inner.w + 6} y={inner.y + inner.h - (inner.h * i) / scales[1].ticks + 3} fontSize={9.5} fill={TEXT_COLOR}>{formatAxisNum(scales[1].min + ((scales[1].max - scales[1].min) * i) / scales[1].ticks)}</text>
              ))}
            </g>
          )}
        </>
      )}
    </g>
  )
}

function WaterfallBody({ inner, data, colors, gridlines, axes, dataLabels }: {
  inner: { x: number; y: number; w: number; h: number }
  data: ExtractedChartData; colors: string[]; gridlines: boolean; axes: boolean; dataLabels: boolean
}) {
  const steps = waterfallSteps(data.series[0]?.values ?? [])
  const n = Math.max(1, steps.length)
  let min = 0, max = 0
  for (const s of steps) { min = Math.min(min, s.start, s.end); max = Math.max(max, s.start, s.end) }
  const { lo, hi, step: unit } = niceRange(min, max)
  const ticks = Math.max(1, Math.round((hi - lo) / unit))
  const yOf = (v: number): number => inner.y + inner.h - ((v - lo) / (hi - lo)) * inner.h
  const slot = inner.w / n
  const barW = Math.max(1, slot * 0.64)
  const shapes: ReactElement[] = []
  steps.forEach((s, i) => {
    const x = inner.x + slot * i + (slot - barW) / 2
    const top = yOf(Math.max(s.start, s.end)), bot = yOf(Math.min(s.start, s.end))
    shapes.push(<rect key={i} x={x} y={top} width={barW} height={Math.max(1, bot - top)} fill={s.kind === 'up' ? colors[0] : colors[1 % colors.length]} />)
    // connector from this bar's end to the next bar
    if (i < steps.length - 1) shapes.push(<line key={'c' + i} x1={x + barW} x2={x + slot} y1={yOf(s.end)} y2={yOf(s.end)} stroke={AXIS_COLOR} strokeWidth={0.75} />)
    if (dataLabels) shapes.push(<text key={'l' + i} x={x + barW / 2} y={top - 3} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(s.end - s.start)}</text>)
  })
  const zeroY = yOf(0)
  return (
    <g>
      {gridlines && Array.from({ length: ticks + 1 }, (_, i) => <line key={i} x1={inner.x} x2={inner.x + inner.w} y1={yOf(lo + unit * i)} y2={yOf(lo + unit * i)} stroke={GRIDLINE_COLOR} strokeWidth={1} />)}
      {shapes}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={zeroY} y2={zeroY} stroke={AXIS_COLOR} />
          {Array.from({ length: ticks + 1 }, (_, i) => (
            <text key={i} x={inner.x - 6} y={yOf(lo + unit * i) + 3} textAnchor="end" fontSize={9.5} fill={TEXT_COLOR}>{formatAxisNum(lo + unit * i)}</text>
          ))}
          <CategoryAxis inner={inner} categories={data.categories} />
        </>
      )}
    </g>
  )
}

function PieBody({ inner, data, colors, doughnut, dataLabels }: {
  inner: { x: number; y: number; w: number; h: number }
  data: ExtractedChartData; colors: string[]; doughnut: boolean; dataLabels: boolean
}) {
  const values = (data.series[0]?.values ?? []).map((v) => Math.max(0, v ?? 0))
  const total = values.reduce((a, b) => a + b, 0) || 1
  const cx = inner.x + inner.w / 2, cy = inner.y + inner.h / 2
  const r = Math.min(inner.w, inner.h) / 2 - 4
  const rInner = doughnut ? r * 0.55 : 0
  let angle = -Math.PI / 2
  const slices: ReactElement[] = []
  const labels: ReactElement[] = []
  values.forEach((v, i) => {
    const frac = v / total
    const a0 = angle
    const a1 = angle + frac * Math.PI * 2
    angle = a1
    const large = a1 - a0 > Math.PI ? 1 : 0
    const p = (a: number, rad: number): [number, number] => [cx + rad * Math.cos(a), cy + rad * Math.sin(a)]
    const [x0, y0] = p(a0, r), [x1, y1] = p(a1, r)
    let d: string
    if (doughnut) {
      const [ix1, iy1] = p(a1, rInner), [ix0, iy0] = p(a0, rInner)
      d = `M${x0} ${y0} A${r} ${r} 0 ${large} 1 ${x1} ${y1} L${ix1} ${iy1} A${rInner} ${rInner} 0 ${large} 0 ${ix0} ${iy0} Z`
    } else {
      d = `M${cx} ${cy} L${x0} ${y0} A${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`
    }
    slices.push(<path key={i} d={d} fill={colors[i % colors.length]} stroke="#FFFFFF" strokeWidth={1} />)
    if (dataLabels && frac > 0.02) {
      const mid = (a0 + a1) / 2
      const [lx, ly] = p(mid, r * 0.68)
      labels.push(
        <text key={'l' + i} x={lx} y={ly} textAnchor="middle" fontSize={9.5} fill="#FFFFFF">
          {Math.round(frac * 100)}%
        </text>
      )
    }
  })
  return <g>{slices}{labels}</g>
}

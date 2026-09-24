// Draws a ChartSpec + its extracted data as inline SVG, styled like an Office 2023+ default chart
// (Aptos font, theme accent colors, light gridlines, "Chart Title" placeholder).
import { ChartSpec } from '../model/types'
import { ExtractedChartData } from './dataExtraction'
import { CHART_FONT, GRIDLINE_COLOR, AXIS_COLOR, TITLE_COLOR, TEXT_COLOR, colorSet } from './theme'

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

function niceMax(v: number): number {
  if (v <= 0) return 1
  const mag = Math.pow(10, Math.floor(Math.log10(v)))
  const norm = v / mag
  const step = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10
  return step * mag
}

export function ChartRenderer({ spec, data, width, height }: ChartRendererProps) {
  const colors = colorSet(spec.colorSetId ?? 0, Math.max(1, data.series.length || 1))
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

function ValueAxis({ inner, max, percent }: { inner: { x: number; y: number; w: number; h: number }; max: number; percent?: boolean }) {
  const ticks = 4
  const labels = []
  for (let i = 0; i <= ticks; i++) {
    const v = (max * i) / ticks
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
  if (Math.abs(v) >= 1000) return (v / 1000).toFixed(v % 1000 === 0 ? 0 : 1) + 'K'
  return Number.isInteger(v) ? String(v) : v.toFixed(1)
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

  let max = 1
  if (stacked) {
    for (let i = 0; i < n; i++) {
      const sum = data.series.reduce((a, s) => a + Math.max(0, s.values[i] ?? 0), 0)
      max = Math.max(max, sum)
    }
  } else {
    for (const s of data.series) for (const v of s.values) max = Math.max(max, v ?? 0)
  }
  max = percent ? 100 : niceMax(max)

  const bars: JSX.Element[] = []
  const catStep = (horizontal ? inner.h : inner.w) / n
  const groupPad = catStep * 0.18
  const groupW = catStep - groupPad * 2
  const barW = stacked ? groupW : groupW / sCount

  for (let i = 0; i < n; i++) {
    let stackAcc = 0
    const total = stacked ? data.series.reduce((a, s) => a + Math.max(0, s.values[i] ?? 0), 0) : 0
    for (let si = 0; si < data.series.length; si++) {
      const raw = data.series[si].values[i] ?? 0
      const val = percent && total > 0 ? (raw / total) * 100 : raw
      const len = (Math.max(0, val) / max) * (horizontal ? inner.w : inner.h)
      if (!horizontal) {
        const x = inner.x + catStep * i + groupPad + (stacked ? 0 : barW * si)
        const y = inner.y + inner.h - len - (stacked ? (stackAcc / max) * inner.h : 0)
        bars.push(<rect key={si + '-' + i} x={x} y={y} width={Math.max(0, barW - 2)} height={Math.max(0, len)} fill={colors[si % colors.length]} />)
        if (dataLabels && len > 10) bars.push(<text key={'l' + si + '-' + i} x={x + barW / 2 - 1} y={y - 3} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(raw)}</text>)
      } else {
        const y = inner.y + catStep * i + groupPad + (stacked ? 0 : barW * si)
        const x = inner.x + (stacked ? (stackAcc / max) * inner.w : 0)
        bars.push(<rect key={si + '-' + i} x={x} y={y} width={Math.max(0, len)} height={Math.max(0, barW - 2)} fill={colors[si % colors.length]} />)
      }
      stackAcc += Math.max(0, val)
    }
  }

  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={4} />}
      {bars}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          {!horizontal && <ValueAxis inner={inner} max={max} percent={percent} />}
          {horizontal && <CategoryAxis inner={inner} categories={data.categories} horizontal />}
          {!horizontal && <CategoryAxis inner={inner} categories={data.categories} />}
        </>
      )}
    </g>
  )
}

function fmtVal(v: number): string {
  return Number.isInteger(v) ? String(v) : v.toFixed(1)
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
  let max = 1
  if (stacked) {
    for (let i = 0; i < n; i++) max = Math.max(max, data.series.reduce((a, s) => a + Math.max(0, s.values[i] ?? 0), 0))
  } else {
    for (const s of data.series) for (const v of s.values) max = Math.max(max, v ?? 0)
  }
  max = niceMax(max)
  const step = n > 1 ? inner.w / (n - 1) : inner.w

  const stackAcc = new Array(n).fill(0)
  const shapes: JSX.Element[] = []
  data.series.forEach((s, si) => {
    const pts: [number, number][] = s.values.map((v, i) => {
      const raw = Math.max(0, v ?? 0)
      const base = stacked ? stackAcc[i] : 0
      if (stacked) stackAcc[i] += raw
      const y = inner.y + inner.h - ((base + raw) / max) * inner.h
      return [inner.x + step * i, y]
    })
    const d = pathFor(pts, false)
    if (area) {
      const baseY = inner.y + inner.h
      const areaD = `${d} L${pts[pts.length - 1][0]} ${baseY} L${pts[0][0]} ${baseY} Z`
      shapes.push(<path key={'a' + si} d={areaD} fill={colors[si % colors.length]} opacity={0.55} stroke="none" />)
    }
    shapes.push(<path key={si} d={d} fill="none" stroke={colors[si % colors.length]} strokeWidth={2} />)
    if (markers) pts.forEach(([x, y], i) => shapes.push(<circle key={si + '-' + i} cx={x} cy={y} r={3} fill="#FFFFFF" stroke={colors[si % colors.length]} strokeWidth={1.5} />))
    if (dataLabels) pts.forEach(([x, y], i) => shapes.push(<text key={'l' + si + '-' + i} x={x} y={y - 6} textAnchor="middle" fontSize={9} fill={TEXT_COLOR}>{fmtVal(s.values[i] ?? 0)}</text>))
  })

  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={4} />}
      {shapes}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <ValueAxis inner={inner} max={max} />
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
  maxY = niceMax(maxY)
  const step = n > 1 ? inner.w / (n - 1) : inner.w

  const shapes: JSX.Element[] = []
  data.series.forEach((s, si) => {
    const pts: [number, number][] = s.values.map((v, i) => [inner.x + step * i, inner.y + inner.h - (Math.max(0, v ?? 0) / maxY) * inner.h])
    if (smooth) shapes.push(<path key={'p' + si} d={pathFor(pts, true)} fill="none" stroke={colors[si % colors.length]} strokeWidth={2} />)
    pts.forEach(([x, y], i) => shapes.push(<circle key={si + '-' + i} cx={x} cy={y} r={3.5} fill={colors[si % colors.length]} />))
  })

  return (
    <g>
      {gridlines && <Gridlines inner={inner} ticks={4} />}
      {shapes}
      {axes && (
        <>
          <line x1={inner.x} x2={inner.x} y1={inner.y} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <line x1={inner.x} x2={inner.x + inner.w} y1={inner.y + inner.h} y2={inner.y + inner.h} stroke={AXIS_COLOR} />
          <ValueAxis inner={inner} max={maxY} />
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
  const slices: JSX.Element[] = []
  const labels: JSX.Element[] = []
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

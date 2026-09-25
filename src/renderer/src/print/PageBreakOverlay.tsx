// View tab > Page Break Preview: draws Excel's blue page-border lines and "Page N" watermarks over
// the normal grid, from the same pagination engine the Print backstage preview uses. Limitation
// (see final report): breaks are shown but not draggable - Excel lets you drag a blue line to move
// a manual break or resize the print area; that interaction isn't implemented here.
import { useMemo } from 'react'
import { useStore, SheetState } from '../model/store'
import { computePrintPlan } from './paginate'
import type { ChartGeometry } from '../charts/ChartOverlay'

export default function PageBreakOverlay({ sheet, geo }: { sheet: SheetState; geo: ChartGeometry }) {
  const store = useStore()
  const { colsAxis, rowsAxis, frozen, scroll, z, rowHeaderW, headerH, viewport } = geo
  const frozenW = colsAxis.totalSize(frozen.cols) * z
  const frozenH = rowsAxis.totalSize(frozen.rows) * z

  const plan = useMemo(() => computePrintPlan(store, sheet), [store, sheet, sheet.layout])

  function screenXRel(c: number): number {
    return c < frozen.cols ? colsAxis.offset(c) * z : frozenW + colsAxis.offset(c) * z - colsAxis.offset(frozen.cols) * z - scroll.x
  }
  function screenYRel(r: number): number {
    return r < frozen.rows ? rowsAxis.offset(r) * z : frozenH + rowsAxis.offset(r) * z - rowsAxis.offset(frozen.rows) * z - scroll.y
  }

  if (plan.empty) return null

  return (
    <div
      className="xl-pagebreak-overlay"
      style={{ position: 'absolute', left: rowHeaderW, top: headerH, width: viewport.w - rowHeaderW, height: viewport.h - headerH, overflow: 'hidden', pointerEvents: 'none' }}
    >
      {plan.pages.map((p, i) => {
        const x0 = screenXRel(p.colStart)
        const x1 = screenXRel(p.colEnd + 1)
        const y0 = screenYRel(p.rowStart)
        const y1 = screenYRel(p.rowEnd + 1)
        const w = x1 - x0, h = y1 - y0
        if (x1 < 0 || y1 < 0 || x0 > viewport.w || y0 > viewport.h) return null
        return (
          <div key={i} style={{ position: 'absolute', left: x0, top: y0, width: Math.max(0, w), height: Math.max(0, h), border: '2px dashed #4472C4', boxSizing: 'border-box' }}>
            <span style={{ position: 'absolute', left: '50%', top: '50%', transform: 'translate(-50%,-50%)', fontSize: 42, fontWeight: 700, color: 'rgba(68,114,196,0.35)', fontFamily: 'Calibri, sans-serif', whiteSpace: 'nowrap' }}>
              Page {p.pageNumber}
            </span>
          </div>
        )
      })}
    </div>
  )
}

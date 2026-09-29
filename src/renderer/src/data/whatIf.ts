// What-If Analysis: Data Table (one- and two-variable) and Scenario Manager.
import { store, SheetState } from '../model/store'
import { Range, ScenarioDef, normRange, colName } from '../model/types'

type Out = { r: number; c: number; raw: string }

const asRaw = (v: unknown): string => {
  if (v === null || v === undefined) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'object' && 'error' in (v as object)) return String((v as { error: unknown }).error)
  return String(v)
}

/**
 * Data > What-If > Data Table over `range` (the block that includes the input row / column).
 * Both inputs: top-left holds the formula, the top row feeds the row input cell, the left column the column input cell.
 * Column input only: values down the left column, formulas across the top row. Row input only: values across the top, formulas down the left.
 * The results are written as values (Excel keeps them live with {=TABLE()}; HyperFormula has no equivalent).
 */
export function computeDataTable(sheet: SheetState, range: Range, rowInput: { r: number; c: number } | null, colInput: { r: number; c: number } | null): Out[] {
  const g = normRange(range)
  const out: Out[] = []
  const restore: Out[] = []
  const remember = (cell: { r: number; c: number } | null): void => { if (cell) restore.push({ r: cell.r, c: cell.c, raw: store.getRaw(sheet, cell.r, cell.c) }) }
  remember(rowInput); remember(colInput)
  const put = (cell: { r: number; c: number } | null, raw: string): void => { if (cell) store.setRaw(sheet, cell.r, cell.c, raw) }
  const read = (r: number, c: number): string => {
    if (store.calcMode === 'manual') store.recalculate()
    return asRaw(store.getValue(sheet, r, c))
  }
  store.transact('Data Table', () => {
    try {
      if (rowInput && colInput) {
        for (let r = g.r1 + 1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) {
          put(rowInput, store.getRaw(sheet, g.r1, c)); put(colInput, store.getRaw(sheet, r, g.c1))
          out.push({ r, c, raw: read(g.r1, g.c1) })
        }
      } else if (colInput) {
        for (let r = g.r1 + 1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) {
          put(colInput, store.getRaw(sheet, r, g.c1))
          out.push({ r, c, raw: read(g.r1, c) })
        }
      } else if (rowInput) {
        for (let r = g.r1 + 1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) {
          put(rowInput, store.getRaw(sheet, g.r1, c))
          out.push({ r, c, raw: read(r, g.c1) })
        }
      }
    } finally {
      for (const x of restore) store.setRaw(sheet, x.r, x.c, x.raw)
    }
    if (out.length) store.setCells(sheet, out, 'Data Table')
  })
  return out
}

// ---------- Scenario Manager ----------
export const scenariosOf = (sheet: SheetState): ScenarioDef[] => sheet.layout.scenarios ?? []

export function cellsOfRanges(ranges: Range[]): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = []
  for (const rg of ranges) {
    const g = normRange(rg)
    for (let r = g.r1; r <= g.r2 && out.length < 32; r++) for (let c = g.c1; c <= g.c2 && out.length < 32; c++) out.push({ r, c })
  }
  return out
}

export function saveScenario(sheet: SheetState, s: ScenarioDef, replace?: string): void {
  store.updateLayout(sheet, replace ? 'Edit Scenario' : 'Add Scenario', (l) => {
    const rest = (l.scenarios ?? []).filter((x) => x.name !== (replace ?? s.name))
    l.scenarios = [...rest, s]
  })
}

export function deleteScenario(sheet: SheetState, name: string): void {
  store.updateLayout(sheet, 'Delete Scenario', (l) => { l.scenarios = (l.scenarios ?? []).filter((x) => x.name !== name) })
}

export function showScenario(sheet: SheetState, s: ScenarioDef): void {
  store.setCells(sheet, s.cells.map((cell, i) => ({ r: cell.r, c: cell.c, raw: s.values[i] ?? '' })), 'Show Scenario')
}

const ref = (r: number, c: number): string => `${colName(c)}${r + 1}`

/** Scenario Summary: a new sheet listing each scenario's values for the changing cells, and the result cells under each. */
export function scenarioSummary(sheet: SheetState, resultCells: { r: number; c: number }[]): void {
  const list = scenariosOf(sheet)
  if (!list.length) return
  const changing: { r: number; c: number }[] = []
  for (const s of list) for (const cell of s.cells) if (!changing.some((x) => x.r === cell.r && x.c === cell.c)) changing.push(cell)
  const original = changing.map((cell) => store.getRaw(sheet, cell.r, cell.c))
  const rows: string[][] = [['Scenario Summary', '', 'Current Values:', ...list.map((s) => s.name)], ['Changing Cells:']]
  changing.forEach((cell, i) => {
    rows.push(['', ref(cell.r, cell.c), original[i], ...list.map((s) => { const k = s.cells.findIndex((x) => x.r === cell.r && x.c === cell.c); return k >= 0 ? s.values[k] : original[i] })])
  })
  if (resultCells.length) {
    rows.push(['Result Cells:'])
    const results = resultCells.map((cell) => ({ cell, vals: [] as string[] }))
    const cellsBefore = changing.map((cell, i) => ({ r: cell.r, c: cell.c, raw: original[i] }))
    store.transact('Scenario Summary', () => {
      results.forEach((x) => x.vals.push(asRaw(store.getValue(sheet, x.cell.r, x.cell.c))))
      for (const s of list) {
        showScenario(sheet, s)
        if (store.calcMode === 'manual') store.recalculate()
        results.forEach((x) => x.vals.push(asRaw(store.getValue(sheet, x.cell.r, x.cell.c))))
      }
      store.setCells(sheet, cellsBefore, 'Scenario Summary')
    })
    for (const x of results) rows.push(['', ref(x.cell.r, x.cell.c), ...x.vals])
  }
  store.addSheet('Scenario Summary')
  const out = store.active
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') store.setRaw(out, r + 1, c + 1, v) }))
  store.setSelection({ ranges: [{ r1: 1, c1: 1, r2: 1, c2: 1 }], active: { r: 1, c: 1 } })
}

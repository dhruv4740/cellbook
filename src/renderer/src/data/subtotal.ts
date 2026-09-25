// Data > Subtotal: inserts a SUBTOTAL row after each group of equal values in one column, plus a
// grand total, like Excel. Remove All takes those rows out again.
import type { WorkbookStore, SheetState } from '../model/store'
import { Range, colName } from '../model/types'

export const SUBTOTAL_FUNCS: { name: string; code: number; label: string }[] = [
  { name: 'Sum', code: 9, label: 'Total' },
  { name: 'Count', code: 3, label: 'Count' },
  { name: 'Average', code: 1, label: 'Average' },
  { name: 'Max', code: 4, label: 'Max' },
  { name: 'Min', code: 5, label: 'Min' },
  { name: 'Product', code: 6, label: 'Product' },
  { name: 'Count Numbers', code: 2, label: 'Count' },
  { name: 'StdDev', code: 7, label: 'StdDev' },
  { name: 'Var', code: 10, label: 'Var' }
]

export interface SubtotalOptions {
  groupCol: number
  func: number // index into SUBTOTAL_FUNCS
  addTo: number[]
  replace: boolean
  summaryBelow: boolean
}

function isSubtotalRow(store: WorkbookStore, sheet: SheetState, r: number, rg: Range): boolean {
  for (let c = rg.c1; c <= rg.c2; c++) if (/^=SUBTOTAL\(/i.test(store.getRaw(sheet, r, c))) return true
  return false
}

/** Removes subtotal rows inside the range; returns the range's new last row. */
export function removeSubtotals(store: WorkbookStore, sheet: SheetState, rg: Range): number {
  let r2 = rg.r2
  store.transact('Remove Subtotals', () => {
    for (let r = rg.r2; r > rg.r1; r--) {
      if (isSubtotalRow(store, sheet, r, rg)) { store.deleteRows(sheet, r, 1); r2-- }
    }
  })
  return r2
}

/** Applies subtotals to rg (row rg.r1 = headers). Returns an error message or null. */
export function applySubtotals(store: WorkbookStore, sheet: SheetState, rg0: Range, o: SubtotalOptions): string | null {
  if (rg0.r2 <= rg0.r1) return 'The range must have a header row and at least one row of data.'
  if (!o.addTo.length) return 'Select at least one column to add a subtotal to.'
  const f = SUBTOTAL_FUNCS[o.func] ?? SUBTOTAL_FUNCS[0]
  store.transact('Subtotal', () => {
    const rg = { ...rg0 }
    if (o.replace) rg.r2 = removeSubtotals(store, sheet, rg)
    const key = (r: number) => store.getDisplay(sheet, r, o.groupCol).text
    // Groups of consecutive equal values: [start, end]
    const groups: [number, number, string][] = []
    let start = rg.r1 + 1
    for (let r = rg.r1 + 2; r <= rg.r2 + 1; r++) {
      if (r > rg.r2 || key(r) !== key(start)) { groups.push([start, r - 1, key(start)]); start = r }
    }
    const first = rg.r1 + 1
    // Insert from the bottom so earlier row numbers stay valid.
    let lastData = rg.r2
    let inserted = 0
    for (let g = groups.length - 1; g >= 0; g--) {
      const [s, e, label] = groups[g]
      const at = o.summaryBelow ? e + 1 : s
      store.insertRows(sheet, at, 1)
      inserted++
      const ds = o.summaryBelow ? s : s + 1, de = o.summaryBelow ? e : e + 1
      const cells = [{ r: at, c: o.groupCol, raw: `${label} ${f.label}` }]
      for (const c of o.addTo) cells.push({ r: at, c, raw: `=SUBTOTAL(${f.code},${colName(c)}${ds + 1}:${colName(c)}${de + 1})` })
      store.setCells(sheet, cells.map((x) => ({ ...x, style: { ...store.getStyle(sheet, x.r, x.c), bold: true } })), 'Subtotal')
    }
    lastData = rg.r2 + inserted
    const gt = o.summaryBelow ? lastData + 1 : first
    store.insertRows(sheet, gt, 1)
    const from = o.summaryBelow ? first : first + 1, to = o.summaryBelow ? lastData : lastData + 1
    const cells = [{ r: gt, c: o.groupCol, raw: `Grand ${f.label}` }]
    for (const c of o.addTo) cells.push({ r: gt, c, raw: `=SUBTOTAL(${f.code},${colName(c)}${from + 1}:${colName(c)}${to + 1})` })
    store.setCells(sheet, cells.map((x) => ({ ...x, style: { ...store.getStyle(sheet, x.r, x.c), bold: true } })), 'Subtotal')
  })
  return null
}

// SUBTOTAL/AGGREGATE need to know which rows are hidden (e.g. by AutoFilter, which writes
// straight into SheetLayout.hiddenRows) so the "ignore hidden rows" function-number variants
// (101-111 for SUBTOTAL, the 4/5/6/7 option codes for AGGREGATE) behave like Excel. HyperFormula
// itself has no notion of row visibility, so the store wires a lookup in here after it is built;
// the plugins (which are plain, store-independent modules registered once at module load) call
// this instead of importing the store directly and creating a cycle.
export type HiddenRowsProvider = (hfSheetId: number) => Record<number, true> | undefined

let provider: HiddenRowsProvider = () => undefined

export function setHiddenRowsProvider(fn: HiddenRowsProvider): void {
  provider = fn
}

export function isRowHidden(hfSheetId: number, row: number): boolean {
  const hidden = provider(hfSheetId)
  return !!hidden && hidden[row] === true
}

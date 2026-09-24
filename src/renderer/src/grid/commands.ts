// Commands that act on the active sheet's current selection. Called from the grid (keyboard,
// right-click menus) and from the ribbon, so both stay consistent.
// STUB signatures: implemented by the grid worker. The ribbon worker may call these freely.

export type PasteKind = 'all' | 'values' | 'formulas' | 'formats' | 'transpose' | 'noBorders' | 'columnWidths' | 'link'

export async function cut(): Promise<void> {}
export async function copy(): Promise<void> {}
export async function paste(kind: PasteKind = 'all'): Promise<void> { void kind }
/** Opens Excel's Insert dialog (Shift cells right / down, Entire row / column) for a cell selection. */
export function insertCellsDialog(): void {}
export function deleteCellsDialog(): void {}
export function insertSheetRows(): void {}
export function insertSheetColumns(): void {}
export function deleteSheetRows(): void {}
export function deleteSheetColumns(): void {}
export function clearSelection(what: 'all' | 'contents' | 'formats'): void { void what }
export function sortSelection(order: 'asc' | 'desc'): void { void order }
/** Fill Down/Right/Up/Left (Ctrl+D / Ctrl+R) over the selection. */
export function fill(direction: 'down' | 'right' | 'up' | 'left'): void { void direction }
/** Home > AutoSum: inserts =SUM(...) (or AVERAGE/COUNT/MAX/MIN) guessing the range like Excel. */
export function autoSum(fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN' = 'SUM'): void { void fn }
export function selectAll(): void {}
/** Moves the active cell to a reference typed in the Name Box ('B7', 'A1:C9'). Returns false if invalid. */
export function goTo(ref: string): boolean { void ref; return false }
export function autoFitColumns(): void {}
export function autoFitRows(): void {}
export function hideRows(hide: boolean): void { void hide }
export function hideColumns(hide: boolean): void { void hide }
export function freezePanes(kind: 'panes' | 'topRow' | 'firstColumn' | 'unfreeze'): void { void kind }
export function mergeCells(kind: 'center' | 'across' | 'merge' | 'unmerge'): void { void kind }

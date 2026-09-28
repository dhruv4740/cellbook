import { shiftFormula } from '../engine/shiftFormula'
// Small UI-only state shared between chrome pieces (not part of the workbook, not undoable).
import { useSyncExternalStore } from 'react'
import { store } from '../model/store'
import { validateInput } from '../data/dataValidation'

export interface UiState {
  backstage: boolean // File tab open
  ribbonTab: string // 'Home', 'Insert', ...
  ribbonCollapsed: boolean
  hideFormulaBar: boolean
  dialog: null | { kind: string; props?: Record<string, unknown> } // e.g. { kind: 'formatCells', props: { tab: 'Number' } }
  /** Sticky format-painter state: set by the Home > Format Painter button, consumed by the grid. */
  formatPainter: null | { once: boolean }
  /**
   * The cell being edited, shared by the in-cell editor (grid) and the formula bar so they stay in sync.
   * null = not editing. Start editing by setting this; commit with commitEdit() / cancel with cancelEdit().
   */
  edit: null | { r: number; c: number; text: string; caret?: number; source: 'cell' | 'bar'; mode: 'Enter' | 'Edit' }
  /** The cell whose Note editor box (Review/right-click > New Note or Edit Note) is open, if any. */
  noteEditor: null | { r: number; c: number }
  /** The cell whose Comment thread card (New Comment, or clicking an existing indicator) is open. */
  commentPanel: null | { r: number; c: number }
  /** The floating chart currently selected on the active sheet, if any; shows the Chart Design/Format tabs. */
  selectedChart: string | null
  /** Ctrl+` (Show Formulas): cells display their raw formula text instead of the computed value. */
  formulasVisible: boolean
  /** F8/Shift+F8: keyboard "Extend Selection" / "Add to Selection" mode (shown on the status bar;
   *  while active, arrow keys extend/add to the selection without needing Shift held down). */
  selectMode: 'none' | 'extend' | 'add'
  /** Which File-backstage pane to land on next time it opens (e.g. Ctrl+F2 -> Print); consumed
   *  once by Backstage.tsx and left as-is otherwise (it defaults to Home). */
  backstageTab: string | null
  /** Scroll Lock: arrow keys scroll the sheet instead of moving the active cell (status bar shows "Scroll Lock"). */
  scrollLock: boolean
  /** "End Mode": armed by a bare End keypress; the next arrow key jumps to the edge of the data
   *  region (like Ctrl+Arrow) once, then disarms (status bar shows "End Mode" while armed). */
  endMode: boolean
  /** Ctrl+Shift+PageUp/PageDown grouped-sheet selection: contiguous sheet indices, or null. */
  groupedSheets: number[] | null
  /** Alt+Down "pick from list": the open dropdown of unique values from the column above/below. */
  pickList: null | { r: number; c: number; x: number; y: number; values: string[] }
  /** Short message shown bottom-centre for a few seconds (e.g. a command Cellbook doesn't have yet). */
  toast: string | null
  /** Ctrl+8: hide the outline symbols (gutter) even when the sheet has groups. */
  outlineHidden: boolean
}

let state: UiState = {
  backstage: false, ribbonTab: 'Home', ribbonCollapsed: false, hideFormulaBar: false, dialog: null,
  formatPainter: null, edit: null, noteEditor: null, commentPanel: null, selectedChart: null,
  formulasVisible: false, selectMode: 'none', backstageTab: null, scrollLock: false, endMode: false,
  groupedSheets: null, pickList: null, toast: null, outlineHidden: false
}
const listeners = new Set<() => void>()

export function setUi(patch: Partial<UiState>): void {
  state = { ...state, ...patch }
  listeners.forEach((l) => l())
}
let toastTimer = 0
/** Tells the user something happened (or didn't) without a modal dialog. */
export function showToast(msg: string): void {
  window.clearTimeout(toastTimer)
  setUi({ toast: msg })
  toastTimer = window.setTimeout(() => setUi({ toast: null }), 4000)
}
/** Used by every command Cellbook doesn't implement yet, so no button silently does nothing. */
export function notAvailable(name: string): void {
  showToast(`"${name.replace(/&/g, '').replace(/\.\.\.$/, '')}" isn't available in Cellbook yet.`)
}
export function getUi(): UiState {
  return state
}
export function useUi(): UiState {
  return useSyncExternalStore((l) => (listeners.add(l), () => listeners.delete(l)), () => state)
}

/** Writes the edit to the active sheet (one undo step) and leaves edit mode. Returns false if nothing was editing. */
export function commitEdit(): boolean {
  const e = state.edit
  if (!e) return false
  const sheet = store.active
  const res = validateInput(store, sheet, e.r, e.c, e.text)
  if (!res.ok) {
    setUi({
      dialog: {
        kind: 'dataValidationError',
        props: {
          style: res.style,
          title: res.title,
          message: res.message,
          onRetry: () => {
            setUi({ dialog: null })
          },
          onCancel: () => {
            cancelEdit()
            setUi({ dialog: null })
          },
          onContinue: () => {
            setUi({ dialog: null, edit: null })
            const raw = store.getRaw(sheet, e.r, e.c)
            if (raw !== e.text) store.setRaw(sheet, e.r, e.c, e.text)
            store.setMode('Ready')
          }
        }
      }
    })
    return false
  }
  setUi({ edit: null })
  const text = normalizeEntry(e.text)
  const raw = store.getRaw(sheet, e.r, e.c)
  if (raw !== text) store.setRaw(sheet, e.r, e.c, text)
  store.setMode('Ready')
  return true
}

/** Excel's entry fix-ups: "+A1"/"-A1" become formulas, and unclosed "(" get their ")". */
export function normalizeEntry(text: string): string {
  let t = text
  if (/^[+-]/.test(t) && !/^[+-]?\s*[\d.,]+%?\s*$/.test(t) && !/^[+-]+$/.test(t)) t = '=' + t
  if (t.startsWith('=') && t.length > 1) {
    let depth = 0
    let inStr = false
    for (const ch of t) {
      if (ch === '"') inStr = !inStr
      else if (!inStr && ch === '(') depth++
      else if (!inStr && ch === ')') depth--
    }
    if (!inStr && depth > 0) t += ')'.repeat(depth)
  }
  return t
}

/** Ctrl+Enter: writes the entry into every selected cell (formulas adjusted per cell), keeps the selection. */
export function commitEditToSelection(): void {
  const e = state.edit
  if (!e) return
  const sheet = store.active
  const text = normalizeEntry(e.text)
  const cells: { r: number; c: number; raw: string }[] = []
  for (const rg of sheet.selection.ranges) {
    const r1 = Math.min(rg.r1, rg.r2), r2 = Math.max(rg.r1, rg.r2)
    const c1 = Math.min(rg.c1, rg.c2), c2 = Math.max(rg.c1, rg.c2)
    if ((r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000) continue
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) cells.push({ r, c, raw: text.startsWith('=') ? shiftFormula(text, r - e.r, c - e.c) : text })
  }
  setUi({ edit: null })
  store.setCells(sheet, cells, 'Typing')
  store.setMode('Ready')
}

export function cancelEdit(): void {
  if (!state.edit) return
  setUi({ edit: null })
  store.setMode('Ready')
}

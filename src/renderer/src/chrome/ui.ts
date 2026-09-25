// Small UI-only state shared between chrome pieces (not part of the workbook, not undoable).
import { useSyncExternalStore } from 'react'
import { store } from '../model/store'
import { validateInput } from '../data/dataValidation'
import { maybeAutoExpandTable, maybeAutoFillCalculatedColumn } from '../model/tables'

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
  /** View tab / status bar workbook view buttons. 'pageBreak' overlays blue page borders and "Page
   *  N" watermarks on the normal grid (see print/PageBreakOverlay.tsx); 'pageLayout' is accepted
   *  and shown as active but doesn't yet change the canvas - see the final report. */
  sheetViewMode: 'normal' | 'pageBreak' | 'pageLayout'
}

let state: UiState = {
  backstage: false, ribbonTab: 'Home', ribbonCollapsed: false, hideFormulaBar: false, dialog: null,
  formatPainter: null, edit: null, noteEditor: null, commentPanel: null, selectedChart: null,
  formulasVisible: false, selectMode: 'none', backstageTab: null, scrollLock: false, endMode: false,
  groupedSheets: null, pickList: null, sheetViewMode: 'normal'
}
const listeners = new Set<() => void>()

export function setUi(patch: Partial<UiState>): void {
  state = { ...state, ...patch }
  listeners.forEach((l) => l())
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
  const raw = store.getRaw(sheet, e.r, e.c)
  if (raw !== e.text) {
    store.transact('Typing', () => {
      store.setRaw(sheet, e.r, e.c, e.text)
      maybeAutoExpandTable(store, sheet, e.r, e.c, e.text)
      maybeAutoFillCalculatedColumn(store, sheet, e.r, e.c, e.text)
    })
  }
  store.setMode('Ready')
  return true
}

export function cancelEdit(): void {
  if (!state.edit) return
  setUi({ edit: null })
  store.setMode('Ready')
}

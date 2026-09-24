// Small UI-only state shared between chrome pieces (not part of the workbook, not undoable).
import { useSyncExternalStore } from 'react'
import { store } from '../model/store'

export interface UiState {
  backstage: boolean // File tab open
  ribbonTab: string // 'Home', 'Insert', ...
  ribbonCollapsed: boolean
  dialog: null | { kind: string; props?: Record<string, unknown> } // e.g. { kind: 'formatCells', props: { tab: 'Number' } }
  /** Sticky format-painter state: set by the Home > Format Painter button, consumed by the grid. */
  formatPainter: null | { once: boolean }
  /**
   * The cell being edited, shared by the in-cell editor (grid) and the formula bar so they stay in sync.
   * null = not editing. Start editing by setting this; commit with commitEdit() / cancel with cancelEdit().
   */
  edit: null | { r: number; c: number; text: string; caret?: number; source: 'cell' | 'bar'; mode: 'Enter' | 'Edit' }
}

let state: UiState = { backstage: false, ribbonTab: 'Home', ribbonCollapsed: false, dialog: null, formatPainter: null, edit: null }
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
  setUi({ edit: null })
  const raw = store.getRaw(store.active, e.r, e.c)
  if (raw !== e.text) store.setRaw(store.active, e.r, e.c, e.text)
  store.setMode('Ready')
  return true
}

export function cancelEdit(): void {
  if (!state.edit) return
  setUi({ edit: null })
  store.setMode('Ready')
}

// Small UI-only state shared between chrome pieces (not part of the workbook, not undoable).
import { useSyncExternalStore } from 'react'

export interface UiState {
  backstage: boolean // File tab open
  ribbonTab: string // 'Home', 'Insert', ...
  ribbonCollapsed: boolean
  dialog: null | { kind: string; props?: Record<string, unknown> } // e.g. { kind: 'formatCells', props: { tab: 'Number' } }
  /** Sticky format-painter state: set by the Home > Format Painter button, consumed by the grid. */
  formatPainter: null | { once: boolean }
}

let state: UiState = { backstage: false, ribbonTab: 'Home', ribbonCollapsed: false, dialog: null, formatPainter: null }
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

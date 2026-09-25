// Store operation synchronization across multiple window views of the same workbook.
// Kept granular (operations, not whole-workbook dumps) so edits stay instantaneous on 500k+ row workbooks.
import type { WorkbookStore, SheetState } from './store'
import type { CellStyle, Range, SheetLayout } from './types'

export type StoreOp =
  | {
      type: 'setRangeRaw'
      sheetName: string
      r: number
      c: number
      raws: string[][]
      label?: string
    }
  | {
      type: 'setCells'
      sheetName: string
      cells: { r: number; c: number; raw: string; style?: CellStyle }[]
      label: string
    }
  | {
      type: 'applyStyle'
      sheetName: string
      ranges: Range[]
      patch: Partial<CellStyle>
      label?: string
    }
  | {
      type: 'setStyles'
      sheetName: string
      styles: Record<string, CellStyle | undefined>
      label: string
    }
  | {
      type: 'updateLayout'
      sheetName: string
      label: string
      layout: SheetLayout
    }
  | {
      type: 'insertRows'
      sheetName: string
      at: number
      count: number
    }
  | {
      type: 'deleteRows'
      sheetName: string
      at: number
      count: number
    }
  | {
      type: 'insertCols'
      sheetName: string
      at: number
      count: number
    }
  | {
      type: 'deleteCols'
      sheetName: string
      at: number
      count: number
    }
  | {
      type: 'addSheet'
      name: string
      index: number
    }
  | {
      type: 'deleteSheet'
      index: number
    }
  | {
      type: 'renameSheet'
      index: number
      name: string
    }
  | {
      type: 'moveSheet'
      from: number
      to: number
    }

export type SyncMessage =
  | { type: 'op'; op: StoreOp }
  | { type: 'transact'; label: string; ops: StoreOp[] }
  | { type: 'undo' }
  | { type: 'redo' }

export function getSheetByName(store: WorkbookStore, name: string): SheetState | undefined {
  return store.sheets.find((s) => s.name === name)
}

export function applyStoreOp(store: WorkbookStore, op: StoreOp): void {
  switch (op.type) {
    case 'setRangeRaw': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.setRangeRaw(sheet, op.r, op.c, op.raws, op.label)
      break
    }
    case 'setCells': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.setCells(sheet, op.cells, op.label)
      break
    }
    case 'applyStyle': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.applyStyle(sheet, op.ranges, op.patch, op.label)
      break
    }
    case 'setStyles': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) {
        const before = new Map<string, CellStyle | undefined>()
        const after = new Map<string, CellStyle | undefined>()
        for (const [k, s] of Object.entries(op.styles)) {
          before.set(k, sheet.styles.get(k))
          after.set(k, s)
        }
        const put = (m: Map<string, CellStyle | undefined>): void =>
          m.forEach((s, k) => {
            const [r, c] = k.split(',').map(Number)
            if (s && Object.keys(s).length) sheet.styles.set(k, s)
            else sheet.styles.delete(k)
          })
        store.commit(op.label, () => put(after), () => put(before))
      }
      break
    }
    case 'updateLayout': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) {
        // Keep each window's own view-specific zoom and freeze panes
        const currentZoom = sheet.layout.zoom
        const currentFrozen = sheet.layout.frozen
        const nextLayout = structuredClone(op.layout)
        nextLayout.zoom = currentZoom
        nextLayout.frozen = currentFrozen
        store.updateLayout(sheet, op.label, (l) => {
          Object.assign(l, nextLayout)
        })
      }
      break
    }
    case 'insertRows': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.insertRows(sheet, op.at, op.count)
      break
    }
    case 'deleteRows': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.deleteRows(sheet, op.at, op.count)
      break
    }
    case 'insertCols': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.insertCols(sheet, op.at, op.count)
      break
    }
    case 'deleteCols': {
      const sheet = getSheetByName(store, op.sheetName)
      if (sheet) store.deleteCols(sheet, op.at, op.count)
      break
    }
    case 'addSheet': {
      const curActive = store.activeIndex
      store.addSheet(op.name, op.index)
      // Preserve this window's independent active sheet
      if (curActive < op.index) {
        store.activeIndex = curActive
      } else {
        store.activeIndex = curActive + 1
      }
      store.notify()
      break
    }
    case 'deleteSheet': {
      const curActive = store.activeIndex
      store.deleteSheet(op.index)
      if (curActive >= store.sheets.length) {
        store.activeIndex = Math.max(0, store.sheets.length - 1)
      }
      store.notify()
      break
    }
    case 'renameSheet': {
      store.renameSheet(op.index, op.name)
      break
    }
    case 'moveSheet': {
      store.moveSheet(op.from, op.to)
      break
    }
  }
}

export function applySyncMessage(store: WorkbookStore, msg: SyncMessage): void {
  if (msg.type === 'undo') {
    store.undo()
  } else if (msg.type === 'redo') {
    store.redo()
  } else if (msg.type === 'op') {
    applyStoreOp(store, msg.op)
  } else if (msg.type === 'transact') {
    store.transact(msg.label, () => {
      for (const op of msg.ops) {
        applyStoreOp(store, op)
      }
    })
  }
}

// Home > Format as Table > New Table Style: user-defined looks, shown under "Custom" in the gallery.
// Format as Table paints plain cell formatting (no table object), so a custom style is just these colours; they are
// remembered per machine (localStorage), not per workbook, and are not written to styles.xml.
export interface TableStyleDef { name: string; header: string; band: string; line: string; headerFont: string }

const KEY = 'cellbook.tableStyles.custom'

/** Adds a style, or replaces the one with the same name (case-insensitive). */
export function upsertStyle(list: TableStyleDef[], s: TableStyleDef): TableStyleDef[] {
  const i = list.findIndex((x) => x.name.toLowerCase() === s.name.toLowerCase())
  return i < 0 ? [...list, s] : list.map((x, k) => (k === i ? s : x))
}

export function loadCustomTableStyles(): TableStyleDef[] {
  try {
    const v: unknown = JSON.parse(localStorage.getItem(KEY) ?? '[]')
    return Array.isArray(v) ? v.filter((s): s is TableStyleDef => !!s && typeof s.name === 'string' && typeof s.header === 'string') : []
  } catch { return [] }
}

export function saveCustomTableStyles(list: TableStyleDef[]): void {
  try { localStorage.setItem(KEY, JSON.stringify(list)) } catch { /* storage unavailable */ }
}

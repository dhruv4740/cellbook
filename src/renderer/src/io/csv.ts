// RFC 4180 CSV read/write.
import type { WorkbookData } from '../model/store'

/** Parses CSV text (comma-delimited, quoted fields, CRLF or LF) into a single-sheet WorkbookData. */
export function readCsv(text: string): WorkbookData {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let inQuotes = false
  let i = 0
  const n = text.length
  const pushField = (): void => { row.push(field); field = '' }
  const pushRow = (): void => { pushField(); rows.push(row); row = [] }

  while (i < n) {
    const ch = text[i]
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue }
        inQuotes = false
        i++
        continue
      }
      field += ch
      i++
      continue
    }
    if (ch === '"') { inQuotes = true; i++; continue }
    if (ch === ',') { pushField(); i++; continue }
    if (ch === '\r') {
      if (text[i + 1] === '\n') i++
      pushRow()
      i++
      continue
    }
    if (ch === '\n') { pushRow(); i++; continue }
    field += ch
    i++
  }
  // last field/row (avoid a trailing empty row if the file ended with a newline)
  if (field.length > 0 || row.length > 0) pushRow()

  const cells: Record<string, { raw?: string }> = {}
  rows.forEach((r, ri) => {
    r.forEach((v, ci) => {
      if (v !== '') cells[ri + ',' + ci] = { raw: v }
    })
  })

  return { sheets: [{ name: 'Sheet1', cells }], activeSheet: 0 }
}

function csvField(v: string): string {
  if (/[",\r\n]/.test(v)) return '"' + v.replace(/"/g, '""') + '"'
  return v
}

/**
 * Writes a rectangular grid of display strings as CSV (CRLF line endings, per RFC 4180).
 * `getDisplay(r, c)` should return the text Excel would show for that cell (empty string past the used range).
 */
export function writeCsv(rows: number, cols: number, getDisplay: (r: number, c: number) => string): string {
  const lines: string[] = []
  for (let r = 0; r < rows; r++) {
    const line: string[] = []
    for (let c = 0; c < cols; c++) line.push(csvField(getDisplay(r, c)))
    lines.push(line.join(','))
  }
  return lines.join('\r\n') + (lines.length ? '\r\n' : '')
}

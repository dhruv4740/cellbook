// Excel number-format engine. STUB: replaced by the number-format worker.
import type { CellValue } from '../model/types'
import { isError } from '../model/types'

export interface FormattedValue {
  text: string
  color?: string // from [Red] etc. in the format code
  align?: 'left' | 'right' | 'center' // natural alignment when hAlign is 'general'
}

/** Formats a computed value with an Excel format code ('General' when undefined). */
export function formatValue(value: CellValue, numFmt?: string): FormattedValue {
  if (value === null || value === undefined) return { text: '' }
  if (isError(value)) return { text: value.error, align: 'center' }
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', align: 'center' }
  if (typeof value === 'number') return { text: String(+value.toPrecision(11)), align: 'right' }
  void numFmt
  return { text: value, align: 'left' }
}

/**
 * Called when the user types into a cell. If the text implies a format (e.g. '$1,200', '15%',
 * '3/4/2025', '1.5E+3'), returns the format code Excel would auto-apply; otherwise undefined.
 */
export function detectInputFormat(raw: string): string | undefined {
  void raw
  return undefined
}

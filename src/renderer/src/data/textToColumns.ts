import { WorkbookStore, SheetState } from '../model/store'
import { Range, normRange, parseA1 } from '../model/types'

export type ColumnFormatType = 'general' | 'text' | 'date' | 'skip'

export interface TextToColumnsOptions {
  mode: 'delimited' | 'fixed'
  delimiters: {
    tab?: boolean
    semicolon?: boolean
    comma?: boolean
    space?: boolean
    other?: boolean
    otherChar?: string
  }
  consecutiveAsOne?: boolean
  textQualifier?: string // '"' | "'" | ''
  breakPositions?: number[] // for fixed width
  columnFormats?: ColumnFormatType[]
  destination?: string // e.g. '$A$1' or 'A1'
}

export function splitDelimitedLine(
  line: string,
  delimiters: string[],
  consecutiveAsOne: boolean,
  textQualifier: string
): string[] {
  if (!delimiters.length) return [line]

  const result: string[] = []
  let cur = ''
  let inQuotes = false

  for (let i = 0; i < line.length; i++) {
    const ch = line[i]

    if (textQualifier && ch === textQualifier) {
      inQuotes = !inQuotes
      continue
    }

    if (!inQuotes && delimiters.includes(ch)) {
      result.push(cur)
      cur = ''
    } else {
      cur += ch
    }
  }
  result.push(cur)

  if (consecutiveAsOne) {
    return result.filter((item, idx) => item !== '' || idx === 0)
  }

  return result
}

export function parseFixedLine(line: string, breakPositions: number[]): string[] {
  const sorted = [...breakPositions].sort((a, b) => a - b)
  const result: string[] = []
  let start = 0

  for (const pos of sorted) {
    if (pos > start && pos < line.length) {
      result.push(line.slice(start, pos).trim())
      start = pos
    }
  }
  result.push(line.slice(start).trim())
  return result
}

export function convertTextToColumns(
  store: WorkbookStore,
  sheet: SheetState,
  range: Range,
  options: TextToColumnsOptions
): void {
  const n = normRange(range)

  // Resolve active delimiters
  const activeDelims: string[] = []
  if (options.delimiters.tab) activeDelims.push('\t')
  if (options.delimiters.semicolon) activeDelims.push(';')
  if (options.delimiters.comma) activeDelims.push(',')
  if (options.delimiters.space) activeDelims.push(' ')
  if (options.delimiters.other && options.delimiters.otherChar) {
    activeDelims.push(options.delimiters.otherChar)
  }

  // Determine destination top-left
  let destR = n.r1
  let destC = n.c1
  if (options.destination) {
    const parsed = parseA1(options.destination)
    if (parsed) {
      destR = parsed.r1
      destC = parsed.c1
    }
  }

  const consecutive = Boolean(options.consecutiveAsOne)
  const qualifier = options.textQualifier || '"'
  const colFormats = options.columnFormats || []

  store.transact('Text to Columns', () => {
    const cells: { r: number; c: number; raw: string }[] = []

    for (let r = n.r1; r <= n.r2; r++) {
      const line = store.getRaw(sheet, r, n.c1)
      let tokens: string[] = []

      if (options.mode === 'fixed' && options.breakPositions && options.breakPositions.length > 0) {
        tokens = parseFixedLine(line, options.breakPositions)
      } else {
        tokens = splitDelimitedLine(line, activeDelims, consecutive, qualifier)
      }

      const rowOffset = r - n.r1

      tokens.forEach((val, colIdx) => {
        const fmt = colFormats[colIdx] || 'general'
        if (fmt === 'skip') return

        const targetR = destR + rowOffset
        const targetC = destC + colIdx

        cells.push({
          r: targetR,
          c: targetC,
          raw: val
        })
      })
    }

    store.setCells(sheet, cells, 'Text to Columns')
  })
}

// Manual column page breaks (<colBreaks>) aren't in exceljs's object model at all (only rowBreaks
// is), so they're patched into the raw worksheet XML after exceljs writes the package - same
// approach as xlsxCfDv.ts/xlsxCharts.ts. OOXML child order requires <colBreaks> right after
// <rowBreaks> (or, if there are none, after <headerFooter>/<pageSetup>/<pageMargins>).
import type { SheetLayout } from '../model/types'

function buildColBreaksXml(colBreaks: number[]): string {
  const sorted = [...new Set(colBreaks)].sort((a, b) => a - b)
  const brks = sorted.map((c) => `<brk id="${c}" max="1048575" man="1"/>`).join('')
  return `<colBreaks count="${sorted.length}" manualBreakCount="${sorted.length}">${brks}</colBreaks>`
}

const INSERT_AFTER = [/<\/rowBreaks>/, /<rowBreaks[^>]*\/>/, /<\/headerFooter>/, /<headerFooter[^>]*\/>/, /<\/pageSetup>/, /<pageSetup[^>]*\/>/, /<pageMargins[^>]*\/>/]

/** Reads `<colBreaks><brk id="c" .../>...</colBreaks>` from a worksheet's raw XML, `c` being the
 *  0-based column index the break falls after - the counterpart to buildColBreaksXml above. */
export function parseColBreaksXml(sheetXml: string): number[] {
  const m = /<colBreaks[^>]*>([\s\S]*?)<\/colBreaks>/.exec(sheetXml)
  if (!m) return []
  return [...m[1].matchAll(/<brk\s+id="(\d+)"/g)].map((b) => Number(b[1]))
}

/** Same, for `<rowBreaks>` - exceljs's plain (non-streaming) Workbook.xlsx.read() doesn't populate
 *  Worksheet.rowBreaks on read even though it writes it, so this is read straight from the XML too. */
export function parseRowBreaksXml(sheetXml: string): number[] {
  const m = /<rowBreaks[^>]*>([\s\S]*?)<\/rowBreaks>/.exec(sheetXml)
  if (!m) return []
  return [...m[1].matchAll(/<brk\s+id="(\d+)"/g)].map((b) => Number(b[1]))
}

export function patchColBreaks(files: Record<string, Uint8Array>, sheetLayouts: (Partial<SheetLayout> | undefined)[]): void {
  sheetLayouts.forEach((layout, i) => {
    const colBreaks = layout?.pageSetup?.colBreaks
    if (!colBreaks || !colBreaks.length) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    const xml = files[path]
    if (!xml) return
    const text = new TextDecoder('utf-8').decode(xml)
    const block = buildColBreaksXml(colBreaks)
    let patched: string | null = null
    for (const re of INSERT_AFTER) {
      if (re.test(text)) { patched = text.replace(re, (m) => m + block); break }
    }
    if (!patched) patched = text.replace('</worksheet>', block + '</worksheet>')
    files[path] = new TextEncoder().encode(patched)
  })
}

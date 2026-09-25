// Excel header/footer field codes: &P page number, &N page count, &D date, &T time, &F file name,
// &A sheet (tab) name, &Z file path, &B/&I/&U toggle bold/italic/underline, &"font,style" and &nn set
// font/size. This resolves the *value* codes to text (bold/italic/font codes are stripped for now -
// the preview/print render uses one font throughout, same limitation noted in the final report).
export interface HFContext {
  page: number
  pageCount: number
  date: Date
  fileName: string
  sheetName: string
  filePath: string
}

export function resolveHFCode(raw: string, ctx: HFContext): string {
  let out = ''
  let i = 0
  while (i < raw.length) {
    const ch = raw[i]
    if (ch === '&' && i + 1 < raw.length) {
      const next = raw[i + 1]
      const upper = next.toUpperCase()
      if (upper === 'P') { out += String(ctx.page); i += 2; continue }
      if (upper === 'N') { out += String(ctx.pageCount); i += 2; continue }
      if (upper === 'D') { out += ctx.date.toLocaleDateString('en-US'); i += 2; continue }
      if (upper === 'T') { out += ctx.date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }); i += 2; continue }
      if (upper === 'F') { out += ctx.fileName; i += 2; continue }
      if (upper === 'A') { out += ctx.sheetName; i += 2; continue }
      if (upper === 'Z') { out += ctx.filePath; i += 2; continue }
      if (next === '&') { out += '&'; i += 2; continue }
      if (upper === 'B' || upper === 'I' || upper === 'U' || upper === 'E') { i += 2; continue } // style toggles: not rendered
      if (next === '"') { // &"FontName,Style" - skip
        const end = raw.indexOf('"', i + 2)
        i = end < 0 ? raw.length : end + 1
        continue
      }
      if (/[0-9]/.test(next)) { // &nn font size - skip the digits
        let j = i + 1
        while (j < raw.length && /[0-9]/.test(raw[j])) j++
        i = j
        continue
      }
    }
    out += ch
    i++
  }
  return out
}

export function resolveHFText(text: { left: string; center: string; right: string }, ctx: HFContext) {
  return {
    left: resolveHFCode(text.left, ctx),
    center: resolveHFCode(text.center, ctx),
    right: resolveHFCode(text.right, ctx)
  }
}

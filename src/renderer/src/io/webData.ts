// Data > From Web: turns a fetched page into importable tables (HTML <table>s, CSV, or a JSON array of records).
export interface WebTable { name: string; rows: string[][] }

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' }
const decode = (s: string): string =>
  s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === '#') return String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10))
    return ENT[e.toLowerCase()] ?? m
  })
const text = (html: string): string => decode(html.replace(/<(script|style)[\s\S]*?<\/\1>/gi, '').replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim()

function htmlTables(html: string): WebTable[] {
  const out: WebTable[] = []
  const re = /<table\b[\s\S]*?<\/table>/gi
  let m: RegExpExecArray | null
  while ((m = re.exec(html))) {
    const t = m[0]
    const cap = /<caption[^>]*>([\s\S]*?)<\/caption>/i.exec(t)
    const rows: string[][] = []
    for (const tr of t.match(/<tr\b[\s\S]*?(?:<\/tr>|(?=<tr\b)|$)/gi) ?? []) {
      const row: string[] = []
      const cre = /<(td|th)\b([^>]*)>([\s\S]*?)(?=<\/t[dh]>|<t[dh]\b|<\/tr>|$)/gi
      let c: RegExpExecArray | null
      while ((c = cre.exec(tr))) {
        const span = Math.max(1, parseInt(/colspan\s*=\s*["']?(\d+)/i.exec(c[2])?.[1] ?? '1', 10))
        row.push(text(c[3]))
        for (let i = 1; i < span; i++) row.push('')
      }
      if (row.length) rows.push(row)
    }
    if (rows.length) out.push({ name: cap ? text(cap[1]) || `Table ${out.length + 1}` : `Table ${out.length + 1}`, rows })
  }
  return out
}

function jsonTable(body: string): WebTable[] {
  let v: unknown
  try { v = JSON.parse(body) } catch { return [] }
  if (!Array.isArray(v) && v && typeof v === 'object') v = Object.values(v).find(Array.isArray) ?? [v]
  if (!Array.isArray(v) || !v.length) return []
  const recs: Record<string, unknown>[] = v.map((x) => (x && typeof x === 'object' ? (x as Record<string, unknown>) : { value: x }))
  const keys = [...new Set(recs.flatMap((r) => Object.keys(r)))]
  const cell = (x: unknown): string => (x == null ? '' : typeof x === 'object' ? JSON.stringify(x) : String(x))
  return [{ name: 'JSON', rows: [keys, ...recs.map((r) => keys.map((k) => cell(r[k])))] }]
}

function csvTable(body: string): WebTable[] {
  const rows: string[][] = []
  let row: string[] = [], f = '', q = false
  for (let i = 0; i < body.length; i++) {
    const ch = body[i]
    if (q) { if (ch === '"') { if (body[i + 1] === '"') { f += '"'; i++ } else q = false } else f += ch }
    else if (ch === '"') q = true
    else if (ch === ',') { row.push(f); f = '' }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && body[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = '' }
    else f += ch
  }
  if (f || row.length) { row.push(f); rows.push(row) }
  return rows.length ? [{ name: 'CSV', rows }] : []
}

export function parseWebTables(body: string, contentType = ''): WebTable[] {
  const ct = contentType.toLowerCase()
  const head = body.trimStart().slice(0, 1)
  if (ct.includes('json') || head === '[' || head === '{') { const j = jsonTable(body); if (j.length) return j }
  if (ct.includes('csv')) return csvTable(body)
  const h = htmlTables(body)
  if (h.length) return h
  return ct.includes('html') ? [] : csvTable(body)
}

// Writes Excel tables (ListObjects) back on save: xl/tables/tableN.xml, its sheet relationship and <tableParts>.
// A table's filter belongs to the table part, so the sheet-level <autoFilter> that covers the same range moves into it.
import type { SheetLayout } from '../model/types'
import { colName } from '../model/types'
import { toFiles, type PkgLike } from './xlsxPostProcess'

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const a1 = (r: number, c: number): string => `${colName(c)}${r + 1}`
const AF = /<autoFilter\b[^>]*\/>|<autoFilter\b[\s\S]*?<\/autoFilter>/

export function injectTables(zipBytes: PkgLike, layouts: (Partial<SheetLayout> | undefined)[], cells: Record<string, { raw?: string }>[]): PkgLike {
  if (!layouts.some((l) => l?.tables?.length)) return zipBytes
  const files = toFiles(zipBytes)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  let ct = dec.decode(files['[Content_Types].xml'])
  let maxTable = 0
  for (const k of Object.keys(files)) { const m = /^xl\/tables\/table(\d+)\.xml$/.exec(k); if (m) maxTable = Math.max(maxTable, +m[1]) }
  const usedNames = new Set<string>()

  layouts.forEach((l, i) => {
    const tables = l?.tables
    if (!tables?.length) return
    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
    if (!files[sheetPath]) return
    let sheetXml = dec.decode(files[sheetPath])
    const relsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let rels = files[relsPath] ? dec.decode(files[relsPath]) : ''
    let maxR = 0
    for (const x of rels.matchAll(/Id="rId(\d+)"/g)) maxR = Math.max(maxR, +x[1])
    const parts: string[] = []
    const relAdd: string[] = []
    for (const t of tables) {
      maxTable++
      let name = t.name.replace(/[^\w.]/g, '_') || `Table${maxTable}`
      if (/^\d/.test(name)) name = '_' + name
      while (usedNames.has(name.toLowerCase())) name += '_'
      usedNames.add(name.toLowerCase())
      const seen = new Set<string>()
      const cols: string[] = []
      for (let c = t.c1; c <= t.c2; c++) {
        let h = t.headerRows ? String(cells[i]?.[`${t.r1},${c}`]?.raw ?? '').replace(/^'/, '') : ''
        if (!h || h.startsWith('=')) h = `Column${c - t.c1 + 1}`
        let n = h, k = 1
        while (seen.has(n.toLowerCase())) n = `${h}${++k}`
        seen.add(n.toLowerCase())
        cols.push(n)
      }
      const filterRef = `${a1(t.r1, t.c1)}:${a1(t.r2 - t.totalsRows, t.c2)}`
      let af = ''
      if (t.filter && t.headerRows) {
        const m = AF.exec(sheetXml)
        if (m && new RegExp(`ref="${filterRef}"`, 'i').test(m[0])) { af = m[0]; sheetXml = sheetXml.replace(m[0], '') } else af = `<autoFilter ref="${filterRef}"/>`
      }
      const style = t.style
        ? `<tableStyleInfo name="${esc(t.style)}" showFirstColumn="${t.firstCol ? 1 : 0}" showLastColumn="${t.lastCol ? 1 : 0}" showRowStripes="${t.rowStripes ? 1 : 0}" showColumnStripes="${t.colStripes ? 1 : 0}"/>`
        : ''
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${maxTable}" name="${esc(name)}" displayName="${esc(name)}" ref="${a1(t.r1, t.c1)}:${a1(t.r2, t.c2)}"` +
        `${t.headerRows ? '' : ' headerRowCount="0"'}${t.totalsRows ? ` totalsRowCount="${t.totalsRows}"` : ' totalsRowShown="0"'}>${af}` +
        `<tableColumns count="${cols.length}">${cols.map((n, k) => `<tableColumn id="${k + 1}" name="${esc(n)}"/>`).join('')}</tableColumns>${style}</table>`
      files[`xl/tables/table${maxTable}.xml`] = enc.encode(xml)
      ct = ct.replace('</Types>', `<Override PartName="/xl/tables/table${maxTable}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`)
      maxR++
      relAdd.push(`<Relationship Id="rId${maxR}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${maxTable}.xml"/>`)
      parts.push(`<tablePart r:id="rId${maxR}"/>`)
    }
    rels = rels.includes('</Relationships>')
      ? rels.replace('</Relationships>', relAdd.join('') + '</Relationships>')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relAdd.join('')}</Relationships>`
    files[relsPath] = enc.encode(rels)
    const frag = `<tableParts count="${parts.length}">${parts.join('')}</tableParts>`
    const re = /<extLst[\s>]|<\/worksheet>/g
    re.lastIndex = Math.max(0, sheetXml.lastIndexOf('</sheetData>'))
    const m = re.exec(sheetXml)
    sheetXml = m ? sheetXml.slice(0, m.index) + frag + sheetXml.slice(m.index) : sheetXml
    files[sheetPath] = enc.encode(sheetXml)
  })
  files['[Content_Types].xml'] = enc.encode(ct)
  return { files }
}

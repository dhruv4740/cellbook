// xl/tables/tableN.xml: Excel Tables (ListObjects) - model/tables.ts's Table[]. exceljs's own
// `ws.addTable` writes a different (older, more limited) shape and doesn't round-trip totalsRowFunction/
// calculatedColumnFormula/tableStyleInfo options cleanly, so tables are written and read the same way
// charts/AutoFilter-criteria are: raw XML parts built/parsed here and spliced into the zip package by
// xlsxPostProcess.ts (write) and fastXlsx.ts/xlsx.ts (read). Framework-free (no exceljs/HyperFormula/
// React) so it can be imported from the worker-side fast reader too.
import type { SheetLayout, Table, TableColumn, TotalsRowFunction } from '../model/types'
import { rangeA1 } from '../model/types'
import { parseRelationships, resolveRelTarget, extractSection, parseAttrs, decodeXmlEntities } from './fastXlsxXml'

const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

const TOTALS_FN_VALUES: TotalsRowFunction[] = ['average', 'count', 'countNums', 'max', 'min', 'sum', 'stdDev', 'var', 'custom']

function buildTableXml(t: Table, tableId: number): string {
  const ref = rangeA1(t.range)
  const dataR2 = t.options.totalRow ? t.range.r2 - 1 : t.range.r2
  const afRef = rangeA1({ r1: t.range.r1, c1: t.range.c1, r2: dataR2, c2: t.range.c2 })
  const cols = t.columns
    .map((col, i) => {
      const attrs = [`id="${i + 1}"`, `name="${escapeXml(col.name)}"`]
      if (t.options.totalRow && col.totalsRowFunction && col.totalsRowFunction !== 'none') attrs.push(`totalsRowFunction="${col.totalsRowFunction}"`)
      if (t.options.totalRow && col.totalsRowLabel) attrs.push(`totalsRowLabel="${escapeXml(col.totalsRowLabel)}"`)
      const body = col.calculatedFormula ? `<calculatedColumnFormula>${escapeXml(col.calculatedFormula.replace(/^=/, ''))}</calculatedColumnFormula>` : ''
      return `<tableColumn ${attrs.join(' ')}>${body}</tableColumn>`
    })
    .join('')
  const si = t.styleName && t.styleName !== 'TableStyleNone'
    ? `<tableStyleInfo name="${escapeXml(t.styleName)}" showFirstColumn="${t.options.firstColumn ? 1 : 0}" showLastColumn="${t.options.lastColumn ? 1 : 0}" showRowStripes="${t.options.bandedRows ? 1 : 0}" showColumnStripes="${t.options.bandedColumns ? 1 : 0}"/>`
    : ''
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tableId}" name="${escapeXml(t.name)}" displayName="${escapeXml(t.name)}" ref="${ref}" totalsRowShown="${t.options.totalRow ? 1 : 0}" headerRowCount="${t.options.headerRow ? 1 : 0}">` +
    (t.options.filterButton ? `<autoFilter ref="${afRef}"/>` : '') +
    `<tableColumns count="${cols.length ? t.columns.length : 0}">${cols}</tableColumns>` +
    si +
    `</table>`
  )
}

/** Patches `files` (the unzipped package exceljs just wrote) in place: one xl/tables/tableN.xml
 *  per Table, a `<tableParts>` element in its sheet's XML, a relationship in that sheet's own
 *  rels, and Content_Types overrides for the new parts. Mirrors injectCharts's shape/ordering. */
export function injectTables(
  files: Record<string, Uint8Array>,
  sheets: { layout?: Partial<SheetLayout> }[]
): void {
  const enc = new TextEncoder()
  const dec = new TextDecoder('utf-8')
  let tableCounter = 0
  const overridesToAdd: string[] = []

  sheets.forEach((sd, i) => {
    const tables = sd.layout?.tables
    if (!tables || !tables.length) return
    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
    const sheetXmlBytes = files[sheetPath]
    if (!sheetXmlBytes) return
    let xml = dec.decode(sheetXmlBytes)

    const relsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let relsXml = files[relsPath]
      ? dec.decode(files[relsPath])
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    let maxId = 0
    for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, parseInt(m[1], 10))

    const tableParts: string[] = []
    tables.forEach((t) => {
      tableCounter++
      const tableName = `table${tableCounter}.xml`
      files[`xl/tables/${tableName}`] = enc.encode(buildTableXml(t, tableCounter))
      overridesToAdd.push(`<Override PartName="/xl/tables/${tableName}" ContentType="${TABLE_CONTENT_TYPE}"/>`)
      maxId++
      const rId = `rId${maxId}`
      relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${rId}" Type="${TABLE_REL_TYPE}" Target="../tables/${tableName}"/></Relationships>`)
      tableParts.push(`<tablePart r:id="${rId}"/>`)
    })
    files[relsPath] = enc.encode(relsXml)

    const el = `<tableParts count="${tableParts.length}">${tableParts.join('')}</tableParts>`
    xml = xml.replace('</worksheet>', el + '</worksheet>')
    files[sheetPath] = enc.encode(xml)
  })

  if (overridesToAdd.length) {
    const ctPath = '[Content_Types].xml'
    const ctXml = dec.decode(files[ctPath])
    files[ctPath] = enc.encode(ctXml.replace('</Types>', overridesToAdd.join('') + '</Types>'))
  }
}

// ---------- read ----------
function parseTableXml(xml: string): Table | null {
  const m = /<table\b([^>]*)>/.exec(xml)
  if (!m) return null
  const a = parseAttrs(m[1], 0, m[1].length)
  const ref = a.ref
  if (!ref) return null
  const [fromA1, toA1] = ref.includes(':') ? ref.split(':') : [ref, ref]
  const range = a1RangeFromRefs(fromA1, toA1)
  if (!range) return null
  const headerRowCount = a.headerRowCount === undefined ? 1 : Number(a.headerRowCount)
  const totalsRowShown = a.totalsRowShown === '1'
  const name = decodeXmlEntities(a.name ?? a.displayName ?? 'Table1')

  const columnsSection = extractSection(xml, 'tableColumns') ?? ''
  const columns: TableColumn[] = []
  const colRe = /<tableColumn\b([^>]*?)(\/?)>/g
  let cm: RegExpExecArray | null
  while ((cm = colRe.exec(columnsSection))) {
    const ca = parseAttrs(cm[1], 0, cm[1].length)
    const col: TableColumn = { name: decodeXmlEntities(ca.name ?? '') }
    const fn = ca.totalsRowFunction as TotalsRowFunction | undefined
    if (fn && TOTALS_FN_VALUES.includes(fn)) col.totalsRowFunction = fn
    if (ca.totalsRowLabel) col.totalsRowLabel = decodeXmlEntities(ca.totalsRowLabel)
    if (cm[2] !== '/') {
      // Has a body - look for <calculatedColumnFormula>...</calculatedColumnFormula> right after this tag.
      const closeIdx = columnsSection.indexOf('</tableColumn>', colRe.lastIndex)
      const body = closeIdx >= 0 ? columnsSection.slice(colRe.lastIndex, closeIdx) : ''
      const fm = /<calculatedColumnFormula[^>]*>([\s\S]*?)<\/calculatedColumnFormula>/.exec(body)
      if (fm) col.calculatedFormula = '=' + decodeXmlEntities(fm[1])
      if (closeIdx >= 0) colRe.lastIndex = closeIdx + '</tableColumn>'.length
    }
    columns.push(col)
  }

  const tsiM = /<tableStyleInfo\b([^>]*)\/?>/.exec(xml)
  const tsi = tsiM ? parseAttrs(tsiM[1], 0, tsiM[1].length) : {}
  const styleName = tsi.name ?? 'TableStyleNone'

  const table: Table = {
    id: 'tbl_' + a.id + '_' + name,
    name,
    range,
    styleName,
    options: {
      headerRow: headerRowCount !== 0,
      totalRow: totalsRowShown,
      bandedRows: tsi.showRowStripes === '1',
      firstColumn: tsi.showFirstColumn === '1',
      lastColumn: tsi.showLastColumn === '1',
      bandedColumns: tsi.showColumnStripes === '1',
      filterButton: /<autoFilter\b/.test(xml)
    },
    columns: columns.length ? columns : Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => ({ name: 'Column' + (i + 1) }))
  }
  return table
}

function colLetterToIndexLocal(s: string): number {
  let n = 0
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}
function a1RangeFromRefs(from: string, to: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const m1 = /^([A-Z]+)(\d+)$/.exec(from)
  const m2 = /^([A-Z]+)(\d+)$/.exec(to)
  if (!m1 || !m2) return null
  return { r1: +m1[2] - 1, c1: colLetterToIndexLocal(m1[1]), r2: +m2[2] - 1, c2: colLetterToIndexLocal(m2[1]) }
}

/** Reads every Table on one sheet, via that sheet's own rels (`<tablePart r:id>` -> xl/tables/tableN.xml). */
export function parseTablesForSheet(files: Record<string, Uint8Array>, dec: TextDecoder, sheetPath: string, sheetXml: string): Table[] {
  if (!/<tableParts\b/.test(sheetXml)) return []
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const file = sheetPath.split('/').pop() ?? ''
  const relsPath = dir + '/_rels/' + file + '.rels'
  const relsBytes = files[relsPath]
  if (!relsBytes) return []
  const relMap = new Map<string, string>()
  for (const r of parseRelationships(dec.decode(relsBytes))) relMap.set(r.id, r.target)
  const tables: Table[] = []
  const partsSection = extractSection(sheetXml, 'tableParts') ?? ''
  const re = /<tablePart\s+r:id="([^"]+)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(partsSection))) {
    const target = relMap.get(m[1])
    if (!target) continue
    const path = resolveRelTarget(sheetPath, target)
    const bytes = files[path]
    if (!bytes) continue
    const t = parseTableXml(dec.decode(bytes))
    if (t) tables.push(t)
  }
  return tables
}

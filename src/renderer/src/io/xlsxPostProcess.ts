// Post-processing on the raw zip package that exceljs's object model can't express: per-column
// AutoFilter criteria (exceljs only round-trips the filtered range, never filterColumn/filters)
// and, for .xlsm saves, the VBA project part exceljs always strips. Runs on the bytes
// wb.xlsx.writeBuffer() already produced, so a normal .xlsx save with no criteria and no macros
// pays nothing beyond the unzip/rezip when there's actually something to patch.
import { unzipSync, zipSync } from 'fflate'
import type { AutoFilter, ChartSpec, SheetLayout } from '../model/types'
import { rangeA1 } from '../model/types'
import { BLANK_TOKEN } from '../engine/filter'
import { parseRelationships, resolveRelTarget } from './fastXlsxXml'
import { buildChartSpaceXml, buildTwoCellAnchorXml } from './xlsxCharts'

export interface VbaProject {
  bin: Uint8Array
  signature?: Uint8Array
  /** Raw bytes of the source workbook's xl/_rels/vbaProject.bin.rels, copied through verbatim -
   *  simplest way to keep the signature part's own relationship intact. */
  signatureRels?: Uint8Array
}

const VBA_REL_TYPE = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Builds the full `<autoFilter ref="..">...</autoFilter>` element, replacing exceljs's
 *  range-only self-closing tag. colId is relative to the range's first column per the OOXML spec. */
function buildAutoFilterXml(af: AutoFilter): string {
  const ref = rangeA1(af.range)
  const body = Object.entries(af.columns)
    .map(([colStr, crit]) => {
      const colId = Number(colStr) - af.range.c1
      if (colId < 0) return ''
      const hasBlank = crit.selected.includes(BLANK_TOKEN)
      const filters = crit.selected
        .filter((v) => v !== BLANK_TOKEN)
        .map((v) => `<filter val="${escapeXmlAttr(v)}"/>`)
        .join('')
      return `<filterColumn colId="${colId}"><filters${hasBlank ? ' blank="1"' : ''}>${filters}</filters></filterColumn>`
    })
    .join('')
  return `<autoFilter ref="${ref}">${body}</autoFilter>`
}

function addContentTypeOverride(ctXml: string, partName: string, contentType: string): string {
  if (ctXml.includes(`PartName="${partName}"`)) return ctXml
  const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`
  return ctXml.replace('</Types>', override + '</Types>')
}

function injectCharts(
  files: Record<string, Uint8Array>,
  sheetCharts: (ChartSpec[] | undefined)[],
  sheetNames: string[],
  sheetCells: Record<string, { raw?: string }>[],
  sheetLayouts?: (Partial<SheetLayout> | undefined)[]
): void {
  const dec = new TextDecoder('utf-8')
  const enc = new TextEncoder()

  const ctPath = '[Content_Types].xml'
  let ctXml = dec.decode(files[ctPath])
  if (!ctXml.includes('Extension="xml"')) {
    ctXml = ctXml.replace('<Types', '<Types><Default Extension="xml" ContentType="application/xml"/>')
  }

  let maxDrawingNum = 0
  let maxChartNum = 0
  for (const key of Object.keys(files)) {
    const mD = /^xl\/drawings\/drawing(\d+)\.xml$/.exec(key)
    if (mD) maxDrawingNum = Math.max(maxDrawingNum, parseInt(mD[1], 10))
    const mC = /^xl\/charts\/chart(\d+)\.xml$/.exec(key)
    if (mC) maxChartNum = Math.max(maxChartNum, parseInt(mC[1], 10))
  }

  for (let i = 0; i < sheetCharts.length; i++) {
    const charts = sheetCharts[i]
    if (!charts || !charts.length) continue

    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
    const sheetBytes = files[sheetPath]
    if (!sheetBytes) continue
    let sheetXml = dec.decode(sheetBytes)
    const sheetName = sheetNames[i] || `Sheet${i + 1}`
    const cells = sheetCells[i] || {}
    const cellReader = {
      get: (r: number, c: number) => {
        const raw = cells[r + ',' + c]?.raw
        if (raw === undefined || raw === '') return null
        if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw)) return Number(raw)
        if (raw === 'TRUE') return true
        if (raw === 'FALSE') return false
        return raw
      }
    }

    const sheetRelsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let sheetRelsXml = files[sheetRelsPath] ? dec.decode(files[sheetRelsPath]) : ''
    let drawingRId = ''
    let drawingPath = ''

    const drawMatch = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)
    if (drawMatch) {
      drawingRId = drawMatch[1]
      const rels = parseRelationships(sheetRelsXml)
      const r = rels.find((rel) => rel.id === drawingRId)
      if (r) drawingPath = resolveRelTarget(sheetPath, r.target)
    }

    if (!drawingPath) {
      maxDrawingNum++
      const drawingFile = `drawing${maxDrawingNum}.xml`
      drawingPath = `xl/drawings/${drawingFile}`

      let maxRId = 0
      for (const m of sheetRelsXml.matchAll(/Id="rId(\d+)"/g)) {
        maxRId = Math.max(maxRId, parseInt(m[1], 10))
      }
      drawingRId = `rId${maxRId + 1}`

      const newRel = `<Relationship Id="${drawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFile}"/>`
      if (sheetRelsXml.includes('</Relationships>')) {
        sheetRelsXml = sheetRelsXml.replace('</Relationships>', newRel + '</Relationships>')
      } else {
        sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${newRel}\n</Relationships>`
      }
      files[sheetRelsPath] = enc.encode(sheetRelsXml)

      const drawingXmlTag = `<drawing r:id="${drawingRId}"/>`
      const insertBeforeRegex = /<(?:tableParts|extLst)[\s\/>]|<\/worksheet>/
      const m = insertBeforeRegex.exec(sheetXml)
      if (m) {
        sheetXml = sheetXml.slice(0, m.index) + drawingXmlTag + sheetXml.slice(m.index)
      } else {
        sheetXml = sheetXml.replace('</worksheet>', drawingXmlTag + '</worksheet>')
      }
      files[sheetPath] = enc.encode(sheetXml)
    }

    const drawingDir = drawingPath.split('/').slice(0, -1).join('/')
    const drawingFileName = drawingPath.split('/').pop() ?? ''
    const drawingRelsPath = `${drawingDir}/_rels/${drawingFileName}.rels`
    let drawingRelsXml = files[drawingRelsPath] ? dec.decode(files[drawingRelsPath]) : ''
    let drawingXml = files[drawingPath] ? dec.decode(files[drawingPath]) : ''

    let maxDrawingRId = 0
    for (const m of drawingRelsXml.matchAll(/Id="rId(\d+)"/g)) {
      maxDrawingRId = Math.max(maxDrawingRId, parseInt(m[1], 10))
    }

    const anchors: string[] = []
    const drawingRelsToAdd: string[] = []

    for (let cIdx = 0; cIdx < charts.length; cIdx++) {
      const spec = charts[cIdx]
      maxChartNum++
      const chartNum = maxChartNum
      const chartPath = `xl/charts/chart${chartNum}.xml`

      maxDrawingRId++
      const chartRelId = `rId${maxDrawingRId}`

      const chartXml = buildChartSpaceXml(spec, sheetName, cellReader)
      files[chartPath] = enc.encode(chartXml)

      const anchorXml = buildTwoCellAnchorXml(spec, chartRelId, cIdx, sheetLayouts?.[i])
      anchors.push(anchorXml)

      drawingRelsToAdd.push(
        `<Relationship Id="${chartRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNum}.xml"/>`
      )

      ctXml = addContentTypeOverride(
        ctXml,
        `/${chartPath}`,
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
      )
    }

    ctXml = addContentTypeOverride(
      ctXml,
      `/${drawingPath}`,
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    )

    if (drawingRelsXml.includes('</Relationships>')) {
      drawingRelsXml = drawingRelsXml.replace(
        '</Relationships>',
        drawingRelsToAdd.join('') + '</Relationships>'
      )
    } else {
      drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${drawingRelsToAdd.join('')}\n</Relationships>`
    }
    files[drawingRelsPath] = enc.encode(drawingRelsXml)

    if (drawingXml.includes('</xdr:wsDr>')) {
      drawingXml = drawingXml.replace('</xdr:wsDr>', anchors.join('') + '</xdr:wsDr>')
    } else {
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">\n${anchors.join('')}\n</xdr:wsDr>`
    }
    files[drawingPath] = enc.encode(drawingXml)
  }

  files[ctPath] = enc.encode(ctXml)
}

function injectVbaProject(files: Record<string, Uint8Array>, vba: VbaProject): void {
  files['xl/vbaProject.bin'] = vba.bin
  if (vba.signature) files['xl/vbaProjectSignature.bin'] = vba.signature
  if (vba.signatureRels) files['xl/_rels/vbaProject.bin.rels'] = vba.signatureRels

  const dec = new TextDecoder('utf-8')
  const enc = new TextEncoder()

  // [Content_Types].xml: a .bin default (covers vbaProject.bin and vbaProjectSignature.bin) plus
  // the macro-enabled content type for the workbook part itself.
  const ctPath = '[Content_Types].xml'
  let ct = dec.decode(files[ctPath])
  if (!/Extension="bin"/.test(ct)) {
    ct = ct.replace(
      '<Default Extension="rels"',
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="rels"'
    )
  }
  ct = ct.replace(
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
    'ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"'
  )
  files[ctPath] = enc.encode(ct)

  // xl/_rels/workbook.xml.rels: point the workbook part at the VBA project part.
  const relsPath = 'xl/_rels/workbook.xml.rels'
  const rels = dec.decode(files[relsPath])
  let maxId = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, parseInt(m[1], 10))
  const rel = `<Relationship Id="rId${maxId + 1}" Type="${VBA_REL_TYPE}" Target="vbaProject.bin"/>`
  files[relsPath] = enc.encode(rels.replace('</Relationships>', rel + '</Relationships>'))
}

/** `sheetFilters[i]` is sheet `i+1`'s AutoFilter (exceljs names worksheet parts sheet1.xml,
 *  sheet2.xml, ... in the order worksheets were added, which matches WorkbookData.sheets order). */
export function postProcessXlsx(
  zipBytes: Uint8Array,
  sheetFilters: (AutoFilter | undefined)[],
  vba?: VbaProject,
  sheetCharts?: (ChartSpec[] | undefined)[],
  sheetNames?: string[],
  sheetCells?: Record<string, { raw?: string }>[],
  sheetLayouts?: (Partial<SheetLayout> | undefined)[]
): Uint8Array {
  const files = unzipSync(zipBytes)
  sheetFilters.forEach((af, i) => {
    if (!af || !Object.keys(af.columns).length) return
    const path = `xl/worksheets/sheet${i + 1}.xml`
    const xml = files[path]
    if (!xml) return
    const text = new TextDecoder('utf-8').decode(xml)
    const patched = text.replace(/<autoFilter ref="[^"]*"\/>/, buildAutoFilterXml(af))
    if (patched !== text) files[path] = new TextEncoder().encode(patched)
  })

  if (sheetCharts && sheetCharts.some((c) => c && c.length > 0)) {
    injectCharts(files, sheetCharts, sheetNames || [], sheetCells || [], sheetLayouts)
  }

  if (vba) injectVbaProject(files, vba)

  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) {
    if (name.endsWith('/')) continue // directory placeholder, not a real entry we need to keep
    out[name] = bytes
  }
  return zipSync(out, { level: 6 })
}

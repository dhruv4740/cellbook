// Post-processing on the raw zip package that exceljs's object model can't express: per-column
// AutoFilter criteria (exceljs only round-trips the filtered range, never filterColumn/filters)
// and, for .xlsm saves, the VBA project part exceljs always strips. Runs on the bytes
// wb.xlsx.writeBuffer() already produced, so a normal .xlsx save with no criteria and no macros
// pays nothing beyond the unzip/rezip when there's actually something to patch.
import { unzipSync, zipSync } from 'fflate'
import type { AutoFilter, DataValidationRule } from '../model/types'
import { rangeA1 } from '../model/types'
import { BLANK_TOKEN } from '../engine/filter'

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

function escapeXmlText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
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

function buildDataValidationXml(rule: DataValidationRule): string {
  const attrs: string[] = []

  if (rule.allow && rule.allow !== 'any') {
    attrs.push(`type="${escapeXmlAttr(rule.allow)}"`)
  }

  if (rule.errorStyle) {
    attrs.push(`errorStyle="${escapeXmlAttr(rule.errorStyle)}"`)
  }

  if (
    rule.operator &&
    rule.allow !== 'list' &&
    rule.allow !== 'custom' &&
    rule.allow !== 'any'
  ) {
    attrs.push(`operator="${escapeXmlAttr(rule.operator)}"`)
  }

  if (rule.ignoreBlank !== undefined) {
    attrs.push(`allowBlank="${rule.ignoreBlank ? '1' : '0'}"`)
  }

  // Note Excel's inverted flag: showDropDown="1" means the in-cell dropdown is HIDDEN.
  if (rule.allow === 'list' && rule.inCellDropdown === false) {
    attrs.push('showDropDown="1"')
  }

  if (rule.showInputMessage !== undefined) {
    attrs.push(`showInputMessage="${rule.showInputMessage ? '1' : '0'}"`)
  }

  if (rule.showErrorAlert !== undefined) {
    attrs.push(`showErrorMessage="${rule.showErrorAlert ? '1' : '0'}"`)
  }

  if (rule.inputTitle) {
    attrs.push(`promptTitle="${escapeXmlAttr(rule.inputTitle)}"`)
  }

  if (rule.inputMessage) {
    attrs.push(`prompt="${escapeXmlAttr(rule.inputMessage)}"`)
  }

  if (rule.errorTitle) {
    attrs.push(`errorTitle="${escapeXmlAttr(rule.errorTitle)}"`)
  }

  if (rule.errorMessage) {
    attrs.push(`error="${escapeXmlAttr(rule.errorMessage)}"`)
  }

  const sqref = rule.ranges.map(rangeA1).join(' ')
  attrs.push(`sqref="${escapeXmlAttr(sqref)}"`)

  let children = ''
  if (rule.formula1 !== undefined && rule.formula1 !== '') {
    let f1 = rule.formula1
    if (rule.allow === 'list' && !f1.startsWith('=') && !f1.startsWith('"')) {
      f1 = `"${f1}"`
    }
    children += `<formula1>${escapeXmlText(f1)}</formula1>`
  }
  if (rule.formula2 !== undefined && rule.formula2 !== '') {
    children += `<formula2>${escapeXmlText(rule.formula2)}</formula2>`
  }

  const attrStr = attrs.length ? ' ' + attrs.join(' ') : ''
  if (children) {
    return `<dataValidation${attrStr}>${children}</dataValidation>`
  }
  return `<dataValidation${attrStr}/>`
}

export function buildDataValidationsXml(rules: DataValidationRule[]): string {
  if (!rules || !rules.length) return ''
  const items = rules.map(buildDataValidationXml).join('')
  return `<dataValidations count="${rules.length}">${items}</dataValidations>`
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
  sheetValidations?: (DataValidationRule[] | undefined)[]
): Uint8Array {
  const files = unzipSync(zipBytes)
  const sheetCount = Math.max(sheetFilters.length, sheetValidations?.length ?? 0)

  for (let i = 0; i < sheetCount; i++) {
    const af = sheetFilters[i]
    const dvs = sheetValidations?.[i]
    const hasAf = af && Object.keys(af.columns).length > 0
    const hasDvs = dvs && dvs.length > 0
    if (!hasAf && !hasDvs) continue

    const path = `xl/worksheets/sheet${i + 1}.xml`
    const xml = files[path]
    if (!xml) continue
    let text = new TextDecoder('utf-8').decode(xml)

    if (hasAf) {
      text = text.replace(/<autoFilter ref="[^"]*"\/>/, buildAutoFilterXml(af!))
    }

    if (hasDvs) {
      const dvXml = buildDataValidationsXml(dvs!)
      if (dvXml) {
        if (text.includes('<dataValidations')) {
          text = text.replace(/<dataValidations[\s\S]*?<\/dataValidations>/, dvXml)
        } else {
          const insertBeforeRegex = /<(?:hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s\/>]|<\/worksheet>/
          const m = insertBeforeRegex.exec(text)
          if (m) {
            text = text.slice(0, m.index) + dvXml + text.slice(m.index)
          } else {
            text = text.replace('</worksheet>', dvXml + '</worksheet>')
          }
        }
      }
    }

    files[path] = new TextEncoder().encode(text)
  }

  if (vba) injectVbaProject(files, vba)

  const out: Record<string, Uint8Array> = {}
  for (const [name, bytes] of Object.entries(files)) {
    if (name.endsWith('/')) continue // directory placeholder, not a real entry we need to keep
    out[name] = bytes
  }
  return zipSync(out, { level: 6 })
}

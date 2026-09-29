// Insert > Pictures in .xlsx: floating images live in a sheet's drawing part (xl/drawings/drawingN.xml, an
// xdr:pic per image, its bytes in xl/media/). Read them into layout.pictures and write them back on save.
import type { PictureDef, SheetLayout } from '../model/types'
import { parseRelationships, resolveRelTarget } from './fastXlsxXml'
import { toFiles, type PkgLike } from './xlsxPostProcess'

const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }
const EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp' }
const EMU = 9525

export function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  let s = ''
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return `data:${mime};base64,${btoa(s)}`
}

export function dataUrlToBytes(src: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(src)
  if (!m) return null
  const bin = atob(m[2])
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return { bytes, mime: m[1] }
}

/** The `xl/media/*` parts the reader should keep (they are binary, so the text-only unzip filter needs to know). */
export const isMediaPart = (name: string): boolean => /^xl\/media\/[^/]+\.(?:png|jpe?g|gif|bmp|webp)$/i.test(name)

const num = (s: string | undefined): number => (s ? parseInt(s, 10) : 0)

export function parseDrawingPictures(
  files: Record<string, Uint8Array>,
  dec: TextDecoder,
  sheetPath: string,
  sheetXml: string,
  colWidths: Record<number, number> = {},
  rowHeights: Record<number, number> = {}
): PictureDef[] {
  const m = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)
  if (!m) return []
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const relsBytes = files[dir + '/_rels/' + (sheetPath.split('/').pop() ?? '') + '.rels']
  if (!relsBytes) return []
  const rel = parseRelationships(dec.decode(relsBytes)).find((r) => r.id === m[1])
  if (!rel) return []
  const drawingPath = resolveRelTarget(sheetPath, rel.target)
  const drawingBytes = files[drawingPath]
  if (!drawingBytes) return []
  const drawingXml = dec.decode(drawingBytes)
  if (!drawingXml.includes('pic>')) return []
  const relBytes = files[drawingPath.split('/').slice(0, -1).join('/') + '/_rels/' + (drawingPath.split('/').pop() ?? '') + '.rels']
  const relMap = new Map<string, string>()
  if (relBytes) for (const r of parseRelationships(dec.decode(relBytes))) relMap.set(r.id, r.target)

  const out: PictureDef[] = []
  const anchorRe = /<(?:xdr:)?(oneCellAnchor|twoCellAnchor)\b[^>]*>([\s\S]*?)<\/(?:xdr:)?\1>/g
  let a: RegExpExecArray | null
  while ((a = anchorRe.exec(drawingXml))) {
    const block = a[2]
    const pic = /<(?:xdr:)?pic\b[\s\S]*?<\/(?:xdr:)?pic>/.exec(block)
    if (!pic) continue
    const embed = /<(?:a:)?blip\b[^>]*r:embed="([^"]+)"/.exec(pic[0])?.[1]
    const target = embed ? relMap.get(embed) : undefined
    if (!target) continue
    const mediaPath = resolveRelTarget(drawingPath, target)
    const bytes = files[mediaPath]
    const mime = MIME[(mediaPath.split('.').pop() ?? '').toLowerCase()]
    if (!bytes || !mime) continue
    const from = /<(?:xdr:)?from>([\s\S]*?)<\/(?:xdr:)?from>/.exec(block)?.[1] ?? ''
    const fc = num(/<(?:xdr:)?col>(\d+)</.exec(from)?.[1]), fr = num(/<(?:xdr:)?row>(\d+)</.exec(from)?.[1])
    const dx = Math.round(num(/<(?:xdr:)?colOff>(-?\d+)</.exec(from)?.[1]) / EMU), dy = Math.round(num(/<(?:xdr:)?rowOff>(-?\d+)</.exec(from)?.[1]) / EMU)
    let width = 0, height = 0
    if (a[1] === 'oneCellAnchor') {
      const ext = /<(?:xdr:)?ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(block)
      width = ext ? Math.round(num(ext[1]) / EMU) : 0
      height = ext ? Math.round(num(ext[2]) / EMU) : 0
    } else {
      const to = /<(?:xdr:)?to>([\s\S]*?)<\/(?:xdr:)?to>/.exec(block)?.[1] ?? ''
      const tc = num(/<(?:xdr:)?col>(\d+)</.exec(to)?.[1]), tr = num(/<(?:xdr:)?row>(\d+)</.exec(to)?.[1])
      const tdx = Math.round(num(/<(?:xdr:)?colOff>(-?\d+)</.exec(to)?.[1]) / EMU), tdy = Math.round(num(/<(?:xdr:)?rowOff>(-?\d+)</.exec(to)?.[1]) / EMU)
      width = tc === fc ? tdx - dx : (colWidths[fc] ?? 64) - dx + tdx
      for (let c = fc + 1; c < tc; c++) width += colWidths[c] ?? 64
      height = tr === fr ? tdy - dy : (rowHeights[fr] ?? 20) - dy + tdy
      for (let r = fr + 1; r < tr; r++) height += rowHeights[r] ?? 20
    }
    if (width <= 0 || height <= 0) continue
    const nm = /<(?:xdr:)?cNvPr\b[^>]*\bname="([^"]*)"/.exec(pic[0])?.[1] ?? `Picture ${out.length + 1}`
    const alt = /<(?:xdr:)?cNvPr\b[^>]*\bdescr="([^"]*)"/.exec(pic[0])?.[1]
    out.push({
      id: 'pic_' + drawingPath.replace(/\D/g, '') + '_' + out.length, name: unescapeXml(nm), src: bytesToDataUrl(bytes, mime),
      anchor: { r: fr, c: fc, dx, dy }, width, height, ...(alt ? { alt: unescapeXml(alt) } : {})
    })
  }
  return out
}

const unescapeXml = (s: string): string => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
const escapeXml = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** Adds each sheet's pictures to its drawing part (creating the drawing if the sheet has none yet). */
export function injectPictures(zipBytes: PkgLike, all: (PictureDef[] | undefined)[], layouts: (Partial<SheetLayout> | undefined)[]): PkgLike {
  void layouts
  if (!all.some((p) => p?.length)) return zipBytes
  const files = toFiles(zipBytes)
  const dec = new TextDecoder('utf-8'), enc = new TextEncoder()
  let ct = dec.decode(files['[Content_Types].xml'])
  let maxDrawing = 0, maxMedia = 0
  for (const k of Object.keys(files)) {
    const d = /^xl\/drawings\/drawing(\d+)\.xml$/.exec(k)
    if (d) maxDrawing = Math.max(maxDrawing, +d[1])
    const md = /^xl\/media\/image(\d+)\./.exec(k)
    if (md) maxMedia = Math.max(maxMedia, +md[1])
  }
  const addDefault = (ext: string, type: string): void => {
    if (new RegExp(`<Default\\b[^>]*Extension="${ext}"`, 'i').test(ct)) return
    ct = ct.replace(/<Types\b[^>]*>/, (m) => m + `<Default Extension="${ext}" ContentType="${type}"/>`)
  }
  const addOverride = (part: string, type: string): void => {
    if (ct.includes(`PartName="${part}"`)) return
    ct = ct.replace('</Types>', `<Override PartName="${part}" ContentType="${type}"/></Types>`)
  }

  all.forEach((pics, i) => {
    if (!pics?.length) return
    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`
    if (!files[sheetPath]) return
    let sheetXml = dec.decode(files[sheetPath])
    const sheetRelsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    let sheetRels = files[sheetRelsPath] ? dec.decode(files[sheetRelsPath]) : ''
    let drawingPath = ''
    const dm = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)
    if (dm) {
      const r = parseRelationships(sheetRels).find((x) => x.id === dm[1])
      if (r) drawingPath = resolveRelTarget(sheetPath, r.target)
    }
    if (!drawingPath) {
      maxDrawing++
      drawingPath = `xl/drawings/drawing${maxDrawing}.xml`
      let maxR = 0
      for (const x of sheetRels.matchAll(/Id="rId(\d+)"/g)) maxR = Math.max(maxR, +x[1])
      const rel = `<Relationship Id="rId${maxR + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${maxDrawing}.xml"/>`
      sheetRels = sheetRels.includes('</Relationships>')
        ? sheetRels.replace('</Relationships>', rel + '</Relationships>')
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`
      files[sheetRelsPath] = enc.encode(sheetRels)
      const tag = `<drawing r:id="rId${maxR + 1}"/>`
      const re = /<(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s/>]|<\/worksheet>/g
      re.lastIndex = Math.max(0, sheetXml.lastIndexOf('</sheetData>'), sheetXml.lastIndexOf('</conditionalFormatting>'))
      const m = re.exec(sheetXml)
      sheetXml = m ? sheetXml.slice(0, m.index) + tag + sheetXml.slice(m.index) : sheetXml.replace('</worksheet>', tag + '</worksheet>')
      files[sheetPath] = enc.encode(sheetXml)
      addOverride('/' + drawingPath, 'application/vnd.openxmlformats-officedocument.drawing+xml')
    }
    const drawingRelsPath = drawingPath.replace(/([^/]+)$/, '_rels/$1.rels')
    let drels = files[drawingRelsPath] ? dec.decode(files[drawingRelsPath]) : ''
    let maxDr = 0
    for (const x of drels.matchAll(/Id="rId(\d+)"/g)) maxDr = Math.max(maxDr, +x[1])
    const relsAdd: string[] = []
    const anchors: string[] = []
    pics.forEach((p, k) => {
      const img = dataUrlToBytes(p.src)
      const ext = img ? EXT[img.mime] : undefined
      if (!img || !ext) return
      maxMedia++
      const mediaName = `image${maxMedia}.${ext}`
      files[`xl/media/${mediaName}`] = img.bytes
      addDefault(ext, img.mime)
      maxDr++
      relsAdd.push(`<Relationship Id="rId${maxDr}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`)
      const cx = Math.round(p.width * EMU), cy = Math.round(p.height * EMU)
      anchors.push(
        `<xdr:oneCellAnchor><xdr:from><xdr:col>${p.anchor.c}</xdr:col><xdr:colOff>${Math.round(p.anchor.dx * EMU)}</xdr:colOff><xdr:row>${p.anchor.r}</xdr:row><xdr:rowOff>${Math.round(p.anchor.dy * EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>` +
        `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${2000 + k}" name="${escapeXml(p.name)}"${p.alt ? ` descr="${escapeXml(p.alt)}"` : ''}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${maxDr}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`
      )
    })
    if (!anchors.length) return
    drels = drels.includes('</Relationships>')
      ? drels.replace('</Relationships>', relsAdd.join('') + '</Relationships>')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsAdd.join('')}</Relationships>`
    files[drawingRelsPath] = enc.encode(drels)
    let dxml = files[drawingPath] ? dec.decode(files[drawingPath]) : ''
    dxml = dxml.includes('</xdr:wsDr>')
      ? dxml.replace('</xdr:wsDr>', anchors.join('') + '</xdr:wsDr>')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors.join('')}</xdr:wsDr>`
    files[drawingPath] = enc.encode(dxml)
  })
  files['[Content_Types].xml'] = enc.encode(ct)
  return { files }
}

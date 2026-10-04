// Insert > Pictures in .xlsx: floating images live in a sheet's drawing part (xl/drawings/drawingN.xml, an
// xdr:pic per image, its bytes in xl/media/). Read them into layout.pictures and write them back on save.
import type { ArrangeProps, PictureDef, ShapeDef, ShapeKind, SheetLayout } from '../model/types'
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

/** Rotation / flips from the object's a:xfrm and the Selection Pane's hidden flag from its cNvPr. */
function arrangeProps(xml: string): ArrangeProps {
  const xf = /<(?:a:)?xfrm\b([^>]*)>/.exec(xml)?.[1] ?? ''
  const rot = Math.round((num(/\brot="(-?\d+)"/.exec(xf)?.[1]) / 60000) % 360 + 360) % 360
  return {
    ...(rot ? { rotation: rot } : {}), ...(/\bflipH="1"/.test(xf) ? { flipH: true } : {}), ...(/\bflipV="1"/.test(xf) ? { flipV: true } : {}),
    ...(/<(?:xdr:)?cNvPr\b[^>]*\bhidden="1"/.test(xml) ? { hidden: true } : {})
  }
}
/** The xfrm attributes and cNvPr hidden attribute for the writer. */
const xfrmAttrs = (o: ArrangeProps): string => `${o.rotation ? ` rot="${Math.round(o.rotation * 60000)}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}`
const hiddenAttr = (o: ArrangeProps): string => (o.hidden ? ' hidden="1"' : '')

export function parseDrawingObjects(
  files: Record<string, Uint8Array>,
  dec: TextDecoder,
  sheetPath: string,
  sheetXml: string,
  colWidths: Record<number, number> = {},
  rowHeights: Record<number, number> = {}
): { pictures: PictureDef[]; shapes: ShapeDef[] } {
  const none = { pictures: [], shapes: [] }
  const m = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)
  if (!m) return none
  const dir = sheetPath.split('/').slice(0, -1).join('/')
  const relsBytes = files[dir + '/_rels/' + (sheetPath.split('/').pop() ?? '') + '.rels']
  if (!relsBytes) return none
  const rel = parseRelationships(dec.decode(relsBytes)).find((r) => r.id === m[1])
  if (!rel) return none
  const drawingPath = resolveRelTarget(sheetPath, rel.target)
  const drawingBytes = files[drawingPath]
  if (!drawingBytes) return none
  const drawingXml = dec.decode(drawingBytes)
  if (!/<(?:xdr:)?(?:pic|sp|cxnSp)\b/.test(drawingXml)) return none
  const relBytes = files[drawingPath.split('/').slice(0, -1).join('/') + '/_rels/' + (drawingPath.split('/').pop() ?? '') + '.rels']
  const relMap = new Map<string, string>()
  if (relBytes) for (const r of parseRelationships(dec.decode(relBytes))) relMap.set(r.id, r.target)

  const pictures: PictureDef[] = [], shapes: ShapeDef[] = []
  const anchorRe = /<(?:xdr:)?(oneCellAnchor|twoCellAnchor)\b[^>]*>([\s\S]*?)<\/(?:xdr:)?\1>/g
  let a: RegExpExecArray | null
  let z = 0 // drawing order = z-order
  while ((a = anchorRe.exec(drawingXml))) {
    const block = a[2]
    const pic = /<(?:xdr:)?pic\b[\s\S]*?<\/(?:xdr:)?pic>/.exec(block)
    const sp = pic ? null : /<(?:xdr:)?(sp|cxnSp)\b[\s\S]*?<\/(?:xdr:)?\1>/.exec(block)
    if (!pic && !sp) continue
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
    const anchor = { r: fr, c: fc, dx, dy }
    if (pic) {
      const embed = /<(?:a:)?blip\b[^>]*r:embed="([^"]+)"/.exec(pic[0])?.[1]
      const target = embed ? relMap.get(embed) : undefined
      if (!target || width <= 0 || height <= 0) continue
      const mediaPath = resolveRelTarget(drawingPath, target)
      const bytes = files[mediaPath]
      const mime = MIME[(mediaPath.split('.').pop() ?? '').toLowerCase()]
      if (!bytes || !mime) continue
      const nm = /<(?:xdr:)?cNvPr\b[^>]*\bname="([^"]*)"/.exec(pic[0])?.[1] ?? `Picture ${pictures.length + 1}`
      const alt = /<(?:xdr:)?cNvPr\b[^>]*\bdescr="([^"]*)"/.exec(pic[0])?.[1]
      pictures.push({ id: 'pic_' + drawingPath.replace(/\D/g, '') + '_' + pictures.length, name: unescapeXml(nm), src: bytesToDataUrl(bytes, mime), anchor, width, height, ...(alt ? { alt: unescapeXml(alt) } : {}), ...arrangeProps(pic[0]), z: z++ })
    } else if (sp) {
      const sh = parseShape(sp[0], sp[1] === 'cxnSp', anchor, width, height)
      if (sh) shapes.push({ ...sh, ...arrangeProps(sp[0]), z: z++, id: 'shp_' + drawingPath.replace(/\D/g, '') + '_' + shapes.length })
    }
  }
  return { pictures, shapes }
}

const PRESETS: Record<string, ShapeKind> = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', rightArrow: 'rightArrow', line: 'line', straightConnector1: 'line' }
const hex = (v: string | undefined): string | undefined => (v ? '#' + v.toLowerCase() : undefined)
const SRGB = /<(?:a:)?solidFill>\s*<(?:a:)?srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/

function parseShape(xml: string, connector: boolean, anchor: ShapeDef['anchor'], width: number, height: number): Omit<ShapeDef, 'id'> | null {
  const spPr = /<(?:xdr:)?spPr\b[\s\S]*?<\/(?:xdr:)?spPr>/.exec(xml)?.[0] ?? ''
  const prst = /<(?:a:)?prstGeom\b[^>]*prst="(\w+)"/.exec(spPr)?.[1] ?? 'rect'
  const txBox = /txBox="1"/.test(xml)
  const ln = /<(?:a:)?ln\b[^>]*>[\s\S]*?<\/(?:a:)?ln>|<(?:a:)?ln\b[^>]*\/>/.exec(spPr)?.[0] ?? ''
  const styled = /<(?:xdr:)?style>/.test(xml)
  let kind: ShapeKind = PRESETS[prst] ?? 'rect'
  const isLine = connector || kind === 'line'
  if (isLine) kind = /<(?:a:)?(?:tail|head)End\b[^>]*type="(?!none)/.test(ln) ? 'arrow' : 'line'
  else if (txBox && kind === 'rect') kind = 'textbox'
  if (width <= 0) return null
  if (height <= 0) height = isLine ? 1 : 0
  if (height <= 0) return null
  const body = ln ? spPr.replace(ln, '') : spPr
  const noFill = /<(?:a:)?noFill\s*\/>/.test(body)
  const fill = isLine || noFill ? undefined : hex(SRGB.exec(body)?.[1]) ?? (styled && kind !== 'textbox' ? '#4472c4' : undefined)
  const line = /<(?:a:)?noFill\s*\/>/.test(ln) ? undefined : hex(SRGB.exec(ln)?.[1]) ?? (styled ? (isLine ? '#4472c4' : '#2f528f') : undefined)
  const lw = /\bw="(\d+)"/.exec(ln)?.[1]
  const paras = [...xml.matchAll(/<(?:a:)?p\b[^>]*>([\s\S]*?)<\/(?:a:)?p>/g)].map((pm) => [...pm[1].matchAll(/<(?:a:)?t\b[^>]*>([\s\S]*?)<\/(?:a:)?t>/g)].map((t) => unescapeXml(t[1])).join(''))
  const text = paras.join('\n').replace(/\n+$/, '')
  const run = /<(?:a:)?r>[\s\S]*?<\/(?:a:)?r>/.exec(xml)?.[0] ?? ''
  const rPr = /<(?:a:)?rPr\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:a:)?rPr>)/.exec(run)
  const sz = rPr ? /\bsz="(\d+)"/.exec(rPr[1])?.[1] : undefined
  const RLN = /<(?:a:)?ln\b[\s\S]*?<\/(?:a:)?ln>/
  const color = rPr?.[2] ? hex(SRGB.exec(rPr[2].replace(RLN, ''))?.[1]) : undefined
  const textOutline = hex(SRGB.exec(RLN.exec(rPr?.[2] ?? '')?.[0] ?? '')?.[1])
  const fontFamily = rPr?.[2] ? /<(?:a:)?latin\b[^>]*typeface="([^"]*)"/.exec(rPr[2])?.[1] : undefined
  const algn = /<(?:a:)?pPr\b[^>]*\balgn="(\w+)"/.exec(xml)?.[1]
  const nm = /<(?:xdr:)?cNvPr\b[^>]*\bname="([^"]*)"/.exec(xml)?.[1] ?? 'Shape'
  const ink = parseInk(spPr, width, height)
  if (ink) {
    const alpha = /<(?:a:)?alpha\b[^>]*val="(\d+)"/.exec(ln)?.[1]
    return { name: unescapeXml(nm), kind: 'ink', anchor, width, height, points: ink, line: line ?? '#000000', lineWidth: lw ? Math.max(0.25, num(lw) / 12700) : 1, ...(alpha ? { alpha: num(alpha) / 100000 } : {}) }
  }
  return {
    name: unescapeXml(nm), kind, anchor, width, height,
    ...(fill ? { fill } : {}), ...(line ? { line, lineWidth: lw ? Math.max(0.25, num(lw) / 12700) : 1 } : {}),
    ...(text ? { text } : {}), ...(color ? { color } : {}),
    ...(sz ? { fontSize: num(sz) / 100 } : {}), ...(rPr && /\bb="1"/.test(rPr[1]) ? { bold: true } : {}), ...(rPr && /\bi="1"/.test(rPr[1]) ? { italic: true } : {}),
    ...(textOutline ? { textOutline } : {}), ...(fontFamily ? { fontFamily: unescapeXml(fontFamily) } : {}), ...(rPr?.[2] && /<(?:a:)?outerShdw\b/.test(rPr[2]) ? { textShadow: true } : {}),
    ...(algn === 'ctr' ? { align: 'ctr' as const } : algn === 'r' ? { align: 'r' as const } : {})
  }
}

/** A freeform (a:custGeom) polyline's points scaled into a width x height px box; null when it has no path. */
function parseInk(spPr: string, width: number, height: number): [number, number][] | null {
  const path = /<(?:a:)?custGeom\b[\s\S]*?<(?:a:)?path\b([^>]*)>([\s\S]*?)<\/(?:a:)?path>/.exec(spPr)
  if (!path) return null
  const pw = num(/\bw="(\d+)"/.exec(path[1])?.[1]) || 1, ph = num(/\bh="(\d+)"/.exec(path[1])?.[1]) || 1
  const pts = [...path[2].matchAll(/<(?:a:)?pt\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/g)].map((m) => [(num(m[1]) / pw) * width, (num(m[2]) / ph) * height] as [number, number])
  return pts.length ? pts : null
}

const unescapeXml = (s: string): string => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
const escapeXml = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

const PRST: Record<ShapeKind, string> = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', rightArrow: 'rightArrow', line: 'line', arrow: 'line', textbox: 'rect', ink: 'rect' }
const rgb = (c: string): string => c.replace('#', '').toUpperCase()

function shapeAnchorXml(s: ShapeDef, id: number): string {
  const cx = Math.round(s.width * EMU), cy = Math.round(s.height * EMU)
  const isLine = s.kind === 'line' || s.kind === 'arrow'
  const from = `<xdr:from><xdr:col>${s.anchor.c}</xdr:col><xdr:colOff>${Math.round(s.anchor.dx * EMU)}</xdr:colOff><xdr:row>${s.anchor.r}</xdr:row><xdr:rowOff>${Math.round(s.anchor.dy * EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>`
  const fill = s.fill ? `<a:solidFill><a:srgbClr val="${rgb(s.fill)}"/></a:solidFill>` : '<a:noFill/>'
  const ln = s.line
    ? `<a:ln w="${Math.round((s.lineWidth ?? 1) * 12700)}"><a:solidFill><a:srgbClr val="${rgb(s.line)}"/></a:solidFill>${s.kind === 'arrow' ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>`
    : '<a:ln><a:noFill/></a:ln>'
  const xfrm = `<a:xfrm${xfrmAttrs(s)}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
  let geom = `<a:prstGeom prst="${PRST[s.kind]}"><a:avLst/></a:prstGeom>`
  let ln2 = ln
  if (s.kind === 'ink') {
    // Freeform polyline; Excel opens it as a freeform shape (it does not rebuild native ink strokes).
    const pts = (s.points ?? []).map((p, i) => `<a:${i ? 'lnTo' : 'moveTo'}><a:pt x="${Math.round(p[0] * EMU)}" y="${Math.round(p[1] * EMU)}"/></a:${i ? 'lnTo' : 'moveTo'}>`).join('')
    geom = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${cx}" h="${cy}" fill="none">${pts}</a:path></a:pathLst></a:custGeom>`
    ln2 = `<a:ln w="${Math.round((s.lineWidth ?? 1) * 12700)}" cap="rnd"><a:solidFill><a:srgbClr val="${rgb(s.line ?? '#000000')}">${s.alpha !== undefined ? `<a:alpha val="${Math.round(s.alpha * 100000)}"/>` : ''}</a:srgbClr></a:solidFill><a:round/></a:ln>`
  }
  const name = escapeXml(s.name)
  if (isLine) {
    return `<xdr:oneCellAnchor>${from}<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"${hiddenAttr(s)}/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>${xfrm}${geom}${ln}</xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:oneCellAnchor>`
  }
  const algn = s.align === 'ctr' ? 'ctr' : s.align === 'r' ? 'r' : 'l'
  const rPr = `<a:rPr lang="en-US" sz="${Math.round((s.fontSize ?? 11) * 100)}"${s.bold ? ' b="1"' : ''}${s.italic ? ' i="1"' : ''}>${s.textOutline ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${rgb(s.textOutline)}"/></a:solidFill></a:ln>` : ''}<a:solidFill><a:srgbClr val="${rgb(s.color ?? '#000000')}"/></a:solidFill>${s.textShadow ? '<a:effectLst><a:outerShdw blurRad="38100" dist="25400" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst>' : ''}${s.fontFamily ? `<a:latin typeface="${escapeXml(s.fontFamily)}"/>` : ''}</a:rPr>`
  const paras = (s.text ?? '').split('\n').map((t) => `<a:p><a:pPr algn="${algn}"/>${t ? `<a:r>${rPr}<a:t>${escapeXml(t)}</a:t></a:r>` : ''}</a:p>`).join('')
  return `<xdr:oneCellAnchor>${from}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"${hiddenAttr(s)}/><xdr:cNvSpPr${s.kind === 'textbox' ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr>${xfrm}${geom}${s.kind === 'ink' ? '<a:noFill/>' : fill}${ln2}</xdr:spPr>` +
    `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="${s.kind === 'textbox' ? 't' : 'ctr'}"/><a:lstStyle/>${paras}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`
}

/** Adds each sheet's pictures and shapes to its drawing part (creating the drawing if the sheet has none yet). */
export function injectPictures(zipBytes: PkgLike, all: (PictureDef[] | undefined)[], layouts: (Partial<SheetLayout> | undefined)[]): PkgLike {
  if (!all.some((p) => p?.length) && !layouts.some((l) => l?.shapes?.length)) return zipBytes
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

  all.forEach((pics0, i) => {
    const pics = pics0 ?? []
    const shapes = layouts[i]?.shapes ?? []
    if (!pics.length && !shapes.length) return
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
    const anchors: { z: number; xml: string }[] = [] // sorted by z below: draw order in the drawing part is the z-order
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
      anchors.push({ z: p.z ?? 1e9 + k, xml:
        `<xdr:oneCellAnchor><xdr:from><xdr:col>${p.anchor.c}</xdr:col><xdr:colOff>${Math.round(p.anchor.dx * EMU)}</xdr:colOff><xdr:row>${p.anchor.r}</xdr:row><xdr:rowOff>${Math.round(p.anchor.dy * EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>` +
        `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${2000 + k}" name="${escapeXml(p.name)}"${hiddenAttr(p)}${p.alt ? ` descr="${escapeXml(p.alt)}"` : ''}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${maxDr}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:xfrm${xfrmAttrs(p)}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`
      })
    })
    shapes.forEach((sh, k) => anchors.push({ z: sh.z ?? 1e9 + 1000 + k, xml: shapeAnchorXml(sh, 3000 + k) }))
    if (!anchors.length) return
    const anchorXml = anchors.sort((x, y) => x.z - y.z).map((x) => x.xml).join('')
    drels = drels.includes('</Relationships>')
      ? drels.replace('</Relationships>', relsAdd.join('') + '</Relationships>')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsAdd.join('')}</Relationships>`
    files[drawingRelsPath] = enc.encode(drels)
    let dxml = files[drawingPath] ? dec.decode(files[drawingPath]) : ''
    dxml = dxml.includes('</xdr:wsDr>')
      ? dxml.replace('</xdr:wsDr>', anchorXml + '</xdr:wsDr>')
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchorXml}</xdr:wsDr>`
    files[drawingPath] = enc.encode(dxml)
  })
  files['[Content_Types].xml'] = enc.encode(ct)
  return { files }
}

// Low-level, allocation-light XML scanning helpers used by fastXlsx.ts to parse worksheet XML
// without a DOMParser. Built on indexOf/charCodeAt string scanning rather than a general XML
// grammar, since spreadsheet XML never nests same-named elements the way we walk them here
// (a <c> never contains another <c>, a <row> never contains another <row>, etc).

export function decodeXmlEntities(s: string): string {
  if (s.indexOf('&') === -1) return s
  let out = ''
  let i = 0
  const n = s.length
  while (i < n) {
    const amp = s.indexOf('&', i)
    if (amp === -1) {
      out += s.slice(i)
      break
    }
    out += s.slice(i, amp)
    const semi = s.indexOf(';', amp + 1)
    if (semi === -1 || semi - amp > 12) {
      out += '&'
      i = amp + 1
      continue
    }
    const ent = s.slice(amp + 1, semi)
    if (ent === 'amp') out += '&'
    else if (ent === 'lt') out += '<'
    else if (ent === 'gt') out += '>'
    else if (ent === 'quot') out += '"'
    else if (ent === 'apos') out += "'"
    else if (ent.charCodeAt(0) === 35 /* '#' */) {
      const isHex = ent.charCodeAt(1) === 120 || ent.charCodeAt(1) === 88
      const code = isHex ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10)
      out += Number.isFinite(code) ? String.fromCodePoint(code) : ''
    } else {
      out += '&' + ent + ';'
    }
    i = semi + 1
  }
  return out
}

function stripPrefix(name: string): string {
  const c = name.indexOf(':')
  return c === -1 ? name : name.slice(c + 1)
}

/** Parses `name="value" name2="value2"` attribute text (namespace prefixes stripped, entities decoded). */
export function parseAttrs(s: string, start: number, end: number): Record<string, string> {
  const attrs: Record<string, string> = {}
  let i = start
  while (i < end) {
    let cc = s.charCodeAt(i)
    while (i < end && (cc === 32 || cc === 9 || cc === 10 || cc === 13 || cc === 47)) {
      i++
      cc = s.charCodeAt(i)
    }
    if (i >= end) break
    const nameStart = i
    while (i < end) {
      cc = s.charCodeAt(i)
      if (cc === 61 /* = */ || cc === 32) break
      i++
    }
    const name = s.slice(nameStart, i)
    while (i < end && s.charCodeAt(i) !== 61) i++
    i++ // skip '='
    while (i < end && s.charCodeAt(i) === 32) i++
    const quote = s.charCodeAt(i)
    if (quote !== 34 && quote !== 39) break // malformed; bail rather than loop forever
    i++
    const valStart = i
    const valEnd = s.indexOf(String.fromCharCode(quote), i)
    if (valEnd === -1) break
    const raw = s.slice(valStart, valEnd)
    if (name) attrs[stripPrefix(name)] = raw.indexOf('&') === -1 ? raw : decodeXmlEntities(raw)
    i = valEnd + 1
  }
  return attrs
}

/** Finds the first `<tag ...>...</tag>` or self-closing `<tag .../>` at or after `from`, returning
 *  its raw attribute text and inner content (empty for self-closing). `tag` must not be a prefix
 *  of another sibling tag name that could appear in the same content (true for all uses here: v/f/is/t). */
export function extractChild(
  s: string,
  tag: string,
  from: number,
  to: number
): { attrs: string; text: string } | null {
  let i = from
  const tagLen = tag.length
  while (true) {
    const idx = s.indexOf('<' + tag, i)
    if (idx === -1 || idx >= to) return null
    const after = idx + 1 + tagLen
    const ch = s.charCodeAt(after)
    if (ch === 62 /* > */ || ch === 32 /* space */ || ch === 47 /* / */) {
      const gt = s.indexOf('>', after)
      if (gt === -1) return null
      const selfClose = s.charCodeAt(gt - 1) === 47
      const attrs = s.slice(after, selfClose ? gt - 1 : gt)
      if (selfClose) return { attrs, text: '' }
      const closeTag = '</' + tag + '>'
      const closeIdx = s.indexOf(closeTag, gt + 1)
      const text = closeIdx === -1 ? s.slice(gt + 1, to) : s.slice(gt + 1, closeIdx)
      return { attrs, text }
    }
    i = idx + 1
  }
}

/** Returns the inner content of the first `<tag>...</tag>` (or '' for `<tag/>`, undefined if absent). */
export function extractSection(xml: string, tag: string): string | undefined {
  const start = xml.indexOf('<' + tag)
  if (start === -1) return undefined
  const after = start + 1 + tag.length
  const ch = xml.charCodeAt(after)
  if (ch !== 62 && ch !== 32 && ch !== 47) return undefined
  const gt = xml.indexOf('>', after)
  if (gt === -1) return undefined
  if (xml.charCodeAt(gt - 1) === 47) return ''
  const close = xml.indexOf('</' + tag + '>', gt + 1)
  return close === -1 ? xml.slice(gt + 1) : xml.slice(gt + 1, close)
}

/** Flattens `<t>text</t>` or `<r><rPr/><t>run</t></r>...` (shared-string / rich-text / inline-string
 *  bodies) into plain text, skipping phonetic `<rPh>` runs. */
export function flattenRichText(inner: string): string {
  if (inner.indexOf('rPh') !== -1) inner = inner.replace(/<rPh[\s\S]*?<\/rPh>/g, '')
  let out = ''
  let i = 0
  const n = inner.length
  while (true) {
    const idx = inner.indexOf('<t', i)
    if (idx === -1 || idx >= n) break
    const after = idx + 2
    const ch = inner.charCodeAt(after)
    if (ch !== 62 && ch !== 32 && ch !== 47) {
      i = idx + 1
      continue
    }
    const gt = inner.indexOf('>', after)
    if (gt === -1) break
    if (inner.charCodeAt(gt - 1) === 47) {
      i = gt + 1
      continue
    }
    const close = inner.indexOf('</t>', gt + 1)
    const text = close === -1 ? inner.slice(gt + 1) : inner.slice(gt + 1, close)
    out += decodeXmlEntities(text)
    i = close === -1 ? inner.length : close + 4
  }
  return out
}

/** Resolves a relationship Target (possibly `../foo.xml`) against the directory of `basePath`. */
export function resolveRelTarget(basePath: string, target: string): string {
  if (target.indexOf('://') !== -1) return target // external target, not a package part
  if (target.charCodeAt(0) === 47 /* '/' */) return target.slice(1)
  const baseDir = basePath.split('/').slice(0, -1)
  const parts = baseDir.concat(target.split('/'))
  const out: string[] = []
  for (const p of parts) {
    if (p === '.' || p === '') continue
    if (p === '..') out.pop()
    else out.push(p)
  }
  return out.join('/')
}

export interface Relationship {
  id: string
  type: string
  target: string
}

export function parseRelationships(xml: string): Relationship[] {
  const out: Relationship[] = []
  const re = /<Relationship\b([^>]*)\/?>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml))) {
    const a = parseAttrs(m[1], 0, m[1].length)
    if (a.Id && a.Target) out.push({ id: a.Id, type: a.Type ?? '', target: a.Target })
  }
  return out
}

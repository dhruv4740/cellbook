import { describe, expect, it } from 'vitest'
import { store } from '../model/store'
import { EQUATIONS, SIGNATURE_TEXT, WORDART_STYLES, insertEquation, insertSignatureLine, insertWordArt } from './textObjects'
import { shapesOf } from './shapes'
import { injectPictures } from '../io/xlsxPictures'

describe('text objects', () => {
  it('WordArt, signature line and equation insert undoable text boxes with their styling', () => {
    const sh = store.active
    insertWordArt(WORDART_STYLES[7])
    insertSignatureLine()
    insertEquation(EQUATIONS[5].text)
    const s = shapesOf(sh)
    expect(s.map((x) => x.kind)).toEqual(['textbox', 'textbox', 'textbox'])
    expect(s[0]).toMatchObject({ text: 'Your text here', textOutline: '#4472c4', textShadow: true, color: '#ffffff' })
    expect(s[1].text).toBe(SIGNATURE_TEXT())
    expect(s[1].text!.split('\n')[0]).toBe('X')
    expect(s[2].text!).toContain('√')
    expect(new Set(s.map((x) => x.id)).size).toBe(3)
    store.undo(); store.undo(); store.undo()
    expect(shapesOf(sh)).toHaveLength(0)
  })
  it('the gallery has distinct, well-formed entries', () => {
    expect(new Set(WORDART_STYLES.map((w) => w.label)).size).toBe(WORDART_STYLES.length)
    for (const w of WORDART_STYLES) expect(w.patch.color).toMatch(/^#[0-9a-f]{6}$/)
    expect(new Set(EQUATIONS.map((q) => q.label)).size).toBe(EQUATIONS.length)
  })
  it('WordArt run properties survive the drawing round trip', async () => {
    const out = injectPictures({ files: { '[Content_Types].xml': new TextEncoder().encode('<Types></Types>'), 'xl/worksheets/sheet1.xml': new TextEncoder().encode('<worksheet><sheetData/></worksheet>') } } as never, [[]] as never, [{ shapes: [{ id: 's', name: 'WordArt 1', kind: 'textbox', anchor: { r: 0, c: 0, dx: 0, dy: 0 }, width: 200, height: 50, text: 'Hi', color: '#ffffff', textOutline: '#4472c4', textShadow: true, fontFamily: 'Impact' }] }] as never)
    const drawing = Object.entries((out as { files: Record<string, Uint8Array> }).files).find(([k]) => /drawings\/drawing\d+\.xml$/.test(k))
    const xml = new TextDecoder().decode(drawing![1])
    expect(xml).toContain('<a:ln w="9525"><a:solidFill><a:srgbClr val="4472C4"/>')
    expect(xml).toContain('<a:outerShdw')
    expect(xml).toContain('<a:latin typeface="Impact"/>')
  })
})

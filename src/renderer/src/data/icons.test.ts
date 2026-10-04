import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { ALL_ICONS, searchIcons, standaloneSvg } from './icons'

describe('icons', () => {
  it('every curated icon renders an SVG and names are unique', () => {
    expect(new Set(ALL_ICONS.map((i) => i.name)).size).toBe(ALL_ICONS.length)
    for (const i of ALL_ICONS) expect(renderToStaticMarkup(createElement(i.Icon))).toMatch(/^<svg\b/)
  })
  it('searches by every word within an optional category', () => {
    expect(searchIcons('', null)).toHaveLength(ALL_ICONS.length)
    expect(searchIcons('arrow up', null).map((i) => i.name)).toContain('Arrow Up')
    expect(searchIcons('arrow', 'Nature')).toHaveLength(0)
    expect(searchIcons('arrow', 'Arrows').length).toBeGreaterThan(5)
    expect(searchIcons('zzzz', null)).toEqual([])
  })
  it('standaloneSvg gives a fixed-size, coloured, namespaced document', () => {
    const svg = standaloneSvg(renderToStaticMarkup(createElement(ALL_ICONS[0].Icon)), '#ff0000', 256)
    expect(svg).toContain('width="256"')
    expect(svg).toContain('height="256"')
    expect(svg).toContain('fill="#ff0000"')
    expect(svg).not.toContain('currentColor')
    expect(svg).not.toContain('1em')
    expect(svg).toContain('xmlns="http://www.w3.org/2000/svg"')
  })
})

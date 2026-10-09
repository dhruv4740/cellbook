// Inserts every new chart family through the Insert tab menus and checks each renders SVG shapes without renderer errors.
import { launch } from './ui_lib.mjs'
const t = await launch('charts2'); const { win } = t
const seed = (rows) => win.evaluate((rows) => { const s = globalThis.store; s.newWorkbook(); const sh = s.active; rows.forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v)))); s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: rows.length - 1, c2: rows[0].length - 1 }], active: { r: 0, c: 0 } }) }, rows)
const click = (label, root = 'button') => win.evaluate(([label, root]) => { const m = (e) => (e.getAttribute('aria-label') || e.title || e.innerText || '').replace(/\s+/g, ' ').trim(); const e = [...document.querySelectorAll(root)].filter((x) => m(x) === label || m(x).startsWith(label)).at(-1); e?.click(); return !!e }, [label, root])
const items = () => win.evaluate(() => [...document.querySelectorAll('[role=menuitem],.xl-menu-item')].map((e) => e.innerText.replace(/\s+/g, ' ').trim()))
const stats = () => win.evaluate(() => { const svg = [...document.querySelectorAll('svg')].filter((s) => s.querySelector('text') && s.getBoundingClientRect().width > 200).at(-1); return svg ? { shapes: svg.querySelectorAll('rect,path,polygon,polyline,circle,line').length, texts: [...svg.querySelectorAll('text')].slice(0, 6).map((x) => x.textContent) } : null })
const data = [['Cat', 'S1', 'S2', 'S3', 'S4'], ['A', 10, 12, 8, 11], ['B', 14, 9, 13, 10], ['C', 7, 15, 9, 14], ['D', 12, 11, 16, 12], ['E', 9, 8, 10, 9], ['F', 15, 14, 12, 13]]
const cases = [
  ['Insert Hierarchy Chart', 'Treemap'], ['Insert Hierarchy Chart', 'Sunburst'],
  ['Insert Statistic Chart', 'Histogram'], ['Insert Statistic Chart', 'Pareto'], ['Insert Statistic Chart', 'Box and Whisker'],
  ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Funnel'], ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'High-Low-Close'],
  ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Open-High-Low-Close'], ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', '3-D Surface'],
  ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Wireframe 3-D Surface'], ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Contour'],
  ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Wireframe Contour'], ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Radar'],
  ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Radar with Markers'], ['Insert Waterfall, Funnel, Stock, Surface or Radar Chart', 'Filled Radar']
]
for (const [btn, item] of cases) {
  await seed(data)
  await t.tab('Insert'); await win.waitForTimeout(700)
  const ok = await click(btn); await win.waitForTimeout(400)
  const okItem = await click(item, '[role=menuitem],.xl-menu-item'); await win.waitForTimeout(700)
  const s = await stats()
  t.log(item, { btn: ok, item: okItem, charts: await win.evaluate(() => globalThis.store.active.layout.charts?.length ?? 0), ...s })
  await t.shot('c2_' + item.replace(/[^a-z0-9]+/gi, '_'))
}
t.log('errors', t.errors)
await t.close()

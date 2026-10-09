// PivotChart, Screenshot dialog and From Web (local http server) - DOM/state checks, no OS screenshots needed.
import http from 'http'
import { launch } from './ui_lib.mjs'
const srv = http.createServer((q, s) => { s.setHeader('content-type', 'text/html'); s.end('<table><caption>Sales</caption><tr><th>Item</th><th>Qty</th></tr><tr><td>apple</td><td>3</td></tr><tr><td>pear</td><td>5</td></tr></table>') }).listen(0)
const port = srv.address().port
const t = await launch('newfeat'); const { win } = t
await win.evaluate(() => {
  const s = globalThis.store; s.newWorkbook(); const sh = s.active
  ;[['Region', 'Sum of Qty'], ['East', 4], ['West', 7], ['Grand Total', 11]].forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v))))
  s.updateLayout(sh, 'x', (l) => { l.pivots = [{ id: 'p1', sourceSheet: sh.name, source: { r1: 0, c1: 0, r2: 3, c2: 1 }, anchor: { r: 0, c: 0 }, rows: [0], cols: [], values: [{ col: 1, func: 'sum' }], lastSize: { rows: 4, cols: 2 } }] })
  s.setSelection({ ranges: [{ r1: 1, c1: 0, r2: 1, c2: 0 }], active: { r: 1, c: 0 } })
})
await t.tab('Insert'); await win.waitForTimeout(700)
const click = (re) => win.evaluate((src) => { const r = new RegExp(src, 'i'); [...document.querySelectorAll('button')].filter((x) => r.test(x.getAttribute('aria-label') || x.title || x.innerText)).at(-1)?.click() }, re)
await click('^PivotChart'); await win.waitForTimeout(700)
t.log('charts', JSON.stringify(await win.evaluate(() => globalThis.store.active.layout.charts?.map((c) => ({ type: c.spec?.type ?? c.type, range: c.spec?.dataRange ?? c.dataRange })) ?? 'none')))
await t.tab('Insert'); await win.waitForTimeout(700); await click('^Screenshot'); await win.waitForTimeout(1200)
t.log('screenshot dialog', JSON.stringify(await win.evaluate(() => ({ dlg: document.body.innerText.includes('Available Windows'), wins: document.querySelectorAll('[data-screenshot-window]').length, clip: [...document.querySelectorAll('button')].some((b) => /Screen Clipping/.test(b.innerText)) }))))
await win.keyboard.press('Escape'); await win.waitForTimeout(300)
await t.tab('Data'); await win.waitForTimeout(700)
await click('^From Web'); await win.waitForTimeout(500)
await win.fill('[data-web-url]', `http://127.0.0.1:${port}/`); await win.keyboard.press('Enter'); await win.waitForTimeout(1500)
t.log('tables', JSON.stringify(await win.evaluate(() => [...document.querySelectorAll('[data-web-table]')].map((e) => e.innerText))))
await win.evaluate(() => [...document.querySelectorAll('button')].find((b) => b.innerText.trim() === 'Load')?.click()); await win.waitForTimeout(600)
t.log('loaded', JSON.stringify(await win.evaluate(() => { const s = globalThis.store; return { sheet: s.active.name, a1: s.getValue(s.active, 0, 0), b3: s.getValue(s.active, 2, 1) } })))
t.log('errors', JSON.stringify(t.errors))
await t.close(); srv.close()

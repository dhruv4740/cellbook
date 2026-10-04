// Forecast Sheet: open the dialog on a 2-column series, create, check the new sheet and chart.
import { launch } from './ui_lib.mjs'
const t = await launch('forecast'); const { win } = t
await win.waitForFunction(() => globalThis.__ui)
await win.evaluate(() => { globalThis.store.newWorkbook() })
for (let i = 0; i < 12; i++) { await t.set(i + 1, 0, String(i + 1)); await t.set(i + 1, 1, String(10 + i * 2 + (i % 3))) }
await t.set(0, 0, 'Period'); await t.set(0, 1, 'Sales')
await t.sel(0, 0, 12, 1)
await t.tab('Data')
await t.click('Forecast Sheet')
await win.waitForTimeout(400)
await t.shot('forecast_dialog')
const title = await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent)
t.log('dialog', title)
const inputs = win.locator('.xl-dialog input[type=text], .xl-dialog input:not([type])')
console.log('inputs', await inputs.count())
await inputs.first().fill('18').catch((e) => console.log('fill fail', String(e).slice(0, 120)))
await win.getByRole('button', { name: /create/i }).click({ force: true })
await win.waitForTimeout(800)
const info = await win.evaluate(() => { const s = globalThis.store; return { sheets: s.sheets.map((x) => x.name), active: s.active.name, c3: s.getValue(s.active, 14, 2), lo: s.getValue(s.active, 14, 3), hi: s.getValue(s.active, 14, 4), raw: s.getRaw(s.active, 14, 2) } })
t.log('after create', info)
await t.shot('forecast_sheet')
t.log('errors', t.errors)
await t.close()

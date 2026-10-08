// Insert > Slicer / Timeline: inserts both, clicks items/periods, checks the filtered rows, clear and remove.
import { launch } from './ui_lib.mjs'
const t = await launch('slicer'); const { win } = t
await win.evaluate(() => {
  const s = globalThis.store; s.newWorkbook(); const sh = s.active
  const rows = [['Fruit', 'When'], ['apple', '2024-01-15'], ['pear', '2024-02-10'], ['apple', '2024-03-05'], ['plum', '2025-01-20']]
  rows.forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, v)))
  s.applyStyle(sh, [{ r1: 1, c1: 1, r2: 4, c2: 1 }], { numFmt: 'yyyy-mm-dd' })
  s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } })
})
const hidden = () => win.evaluate(() => Object.keys(globalThis.store.active.layout.hiddenRows).map(Number).sort())
await t.tab('Insert')
await win.locator('button', { hasText: /^Slicer$/ }).last().evaluate((e) => e.click()); await win.waitForTimeout(400)
t.log('slicer count', await win.evaluate(() => globalThis.store.active.layout.slicers?.length))
t.log('items', await win.evaluate(() => [...document.querySelectorAll('[data-slicer-item]')].map((e) => e.textContent)))
await win.locator('[data-slicer-item="apple"]').click(); await win.waitForTimeout(250)
t.log('apple -> hidden (expect 2,4)', await hidden())
await win.locator('[data-slicer-item="pear"]').click({ modifiers: ['Control'] }); await win.waitForTimeout(250)
t.log('apple+pear -> hidden (expect 4)', await hidden())
await win.locator('button[aria-label="Clear Filter"]').last().click(); await win.waitForTimeout(250)
t.log('cleared', await hidden())
await win.evaluate(() => { const s = globalThis.store; s.setSelection({ ranges: [{ r1: 1, c1: 1, r2: 1, c2: 1 }], active: { r: 1, c: 1 } }) })
await win.locator('button', { hasText: /^Timeline$/ }).last().evaluate((e) => e.click()); await win.waitForTimeout(400)
t.log('periods', await win.evaluate(() => [...document.querySelectorAll('[data-timeline-period]')].map((e) => e.textContent).slice(0, 5)))
await win.locator('[data-timeline-period="Feb 2024"]').click(); await win.waitForTimeout(250)
t.log('Feb 2024 -> hidden (expect 1,3,4)', await hidden())
await win.locator('[data-timeline-period="Mar 2024"]').click({ modifiers: ['Shift'] }); await win.waitForTimeout(250)
t.log('Feb-Mar -> hidden (expect 1,4)', await hidden())
await t.shot('slicer')
await win.evaluate(() => { const s = globalThis.store; for (const x of [...s.active.layout.slicers]) { /* via UI delete */ } })
await win.locator('.xl-slicer').first().click({ button: 'right', force: true }); await win.waitForTimeout(200)
await win.getByText(/^Remove (Slicer|Timeline)$/).first().click({ force: true }); await win.waitForTimeout(250)
t.log('after remove', await win.evaluate(() => globalThis.store.active.layout.slicers?.length), await hidden())
t.log('errors', t.errors); await t.close()

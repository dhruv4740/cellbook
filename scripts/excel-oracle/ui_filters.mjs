// AutoFilter dropdowns: Text / Number / Date filter submenus, applied on real data; checks which rows end up hidden.
import { launch } from './ui_lib.mjs'
const t = await launch('filters'); const { win } = t
const today = Math.floor((Date.now() - new Date().getTimezoneOffset() * 60000) / 86400000) + 25569
await win.evaluate((today) => {
  const s = globalThis.store; s.newWorkbook(); const sh = s.active
  const rows = [['Name', 'Qty', 'When'], ['apple', 3, today], ['pear', 5, today - 1], ['plum', 12, today + 1], ['fig', 8, today - 400], ['kiwi', 1, today + 40]]
  rows.forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v))))
  s.applyStyle(sh, [{ r1: 1, c1: 2, r2: 5, c2: 2 }], { numFmt: 'dd-mm-yyyy' })
  s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 5, c2: 2 }], active: { r: 0, c: 0 } }); s.toggleAutoFilter()
}, today)
const hidden = () => win.evaluate(() => Object.keys(globalThis.store.active.layout.hiddenRows).map(Number).sort())
const open = async (x) => { await win.mouse.click(x, 213); await win.waitForTimeout(250); return win.evaluate(() => !!document.querySelector('.xl-afmenu')) }
const closeMenu = async () => { await win.keyboard.press('Escape'); await win.evaluate(() => document.querySelector('.xl-afmenu-scrim')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))); await win.waitForTimeout(150) }
const item = (txt) => win.locator('.xl-afmenu-item, .xl-afmenu-sub .xl-afmenu-item').filter({ hasText: txt }).first()
const hoverSub = async (txt) => { await item(txt).hover(); await win.waitForTimeout(200) }
t.log('hidden initially', await hidden())
t.log('date menu opens', await open(229))
t.log('menu items', await win.evaluate(() => [...document.querySelectorAll('.xl-afmenu > .xl-afmenu-item')].map((e) => e.innerText.replace(/\s+/g, ' ').trim())))
await hoverSub('Date Filters')
t.log('date submenu', await win.evaluate(() => [...document.querySelectorAll('.xl-afmenu-sub .xl-afmenu-item')].map((e) => e.innerText.trim())))
await win.locator('.xl-afmenu-sub .xl-afmenu-item', { hasText: /^Today$/ }).click({ force: true }); await win.waitForTimeout(250)
t.log('Today -> hidden rows (expect 1,2,3,4 -> idx 2,3,4,5)', await hidden())
await open(229); await hoverSub('Date Filters'); await win.locator('.xl-afmenu-sub .xl-afmenu-item', { hasText: /^This Year$/ }).click({ force: true }); await win.waitForTimeout(250)
t.log('This Year -> hidden', await hidden())
await open(229); await win.locator('.xl-afmenu-item', { hasText: /^Clear Filter/ }).click({ force: true }); await win.waitForTimeout(200)
t.log('cleared', await hidden())
await open(165); await hoverSub('Number Filters'); await win.locator('.xl-afmenu-sub .xl-afmenu-item', { hasText: 'Above Average' }).click({ force: true }); await win.waitForTimeout(250)
t.log('Number > avg(5.8) -> hidden', await hidden())
await open(165); await win.locator('.xl-afmenu-item', { hasText: /^Clear Filter/ }).click({ force: true }); await win.waitForTimeout(200)
await open(101); await hoverSub('Text Filters'); await win.locator('.xl-afmenu-sub .xl-afmenu-item', { hasText: 'Contains' }).first().click({ force: true }); await win.waitForTimeout(250)
t.log('dialog', await win.evaluate(() => document.querySelector('.xl-dialog')?.innerText.replace(/\s+/g,' ')))
await win.locator('.xl-dialog input').first().fill('p'); t.log('fields', await win.evaluate(() => [document.querySelector('.xl-dialog select')?.value, document.querySelector('.xl-dialog input')?.value]))
await win.getByRole('button', { name: 'OK', exact: true }).last().click({ force: true }); await win.waitForTimeout(300)
t.log('af', await win.evaluate(() => JSON.stringify(globalThis.store.active.layout.autoFilter)))
t.log('Text contains p -> hidden (fig,kiwi = 4,5)', await hidden())
await win.evaluate(() => { const s = globalThis.store; s.clearAllFilters?.(s.active); s.applyStyle(s.active, [{ r1: 3, c1: 0, r2: 3, c2: 0 }], { fill: '#ffff00' }) })
await open(101); await hoverSub('Filter by Color')
t.log('color menu', await win.evaluate(() => [...document.querySelectorAll('.xl-afmenu-sub .xl-afmenu-item')].map((e) => e.innerText.trim())))
await win.locator('.xl-afmenu-sub .xl-afmenu-item', { hasText: '#ffff00' }).click({ force: true }); await win.waitForTimeout(250)
t.log('Filter by yellow fill -> hidden (all but plum idx3)', await hidden())
t.log('errors', t.errors); await t.shot('filters'); await t.close()

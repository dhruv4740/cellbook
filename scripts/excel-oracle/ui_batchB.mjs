// Visual check of Quick Analysis, Combo/Waterfall charts, Themes, Icons, WordArt and Sheet Background.
import { launch } from './ui_lib.mjs'
const t = await launch('batchb'); const { win } = t
const btn = (txt) => win.locator('.xl-ribbon-groups button').filter({ hasText: txt }).or(win.locator(`.xl-ribbon-groups button[aria-label*="${txt}"]`)).last()
const go = async (tab, txt) => { await t.tab(tab); await btn(txt).evaluate((e) => e.click()); await win.waitForTimeout(350) }
const menu = async (txt) => { await win.locator('.xl-menu .xl-menu-item, .xl-combo-list .xl-combo-item').filter({ hasText: txt }).first().evaluate((e) => e.click()); await win.waitForTimeout(400) }
await win.evaluate(() => { const s = globalThis.store; s.newWorkbook(); const sh = s.active; [['Item', 'Sales', 'Cost'], ['A', 100, 60], ['B', -30, 20], ['C', 50, 30], ['D', 80, 45], ['E', -20, 10]].forEach((r, i) => r.forEach((v, c) => s.setRaw(sh, i, c, String(v)))); s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 5, c2: 2 }], active: { r: 0, c: 0 } }) })
const step = async (name, fn) => { try { await fn(); await t.shot(name); t.log(name, 'ok') } catch (e) { t.log(name, String(e).slice(0, 200)) } }
await step('qa', async () => { await win.mouse.click(141, 252, { button: 'right' }); await win.waitForTimeout(250); await menu('Quick Analysis') })
await win.keyboard.press('Escape')
await step('combo', async () => { await go('Insert', 'Combo Chart'); await menu('Clustered Column') })
await step('waterfall', async () => { await win.evaluate(() => globalThis.__ui.setUi({ dialog: null })); await go('Insert', 'Waterfall, Funnel'); await menu('Waterfall') })
await step('themes', async () => { await go('Page Layout', 'Colors'); await menu('Blue'); })
await step('icons', async () => { await win.keyboard.press('Escape'); await go('Insert', 'Icons') })
await win.keyboard.press('Escape'); await win.evaluate(() => globalThis.__ui.setUi({ dialog: null }))
await step('wordart', async () => { await go('Insert', 'WordArt'); await menu('Fill') })
t.log('errors', t.errors); await t.close()

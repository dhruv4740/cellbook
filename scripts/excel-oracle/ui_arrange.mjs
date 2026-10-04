// Arrange: insert 3 shapes, rotate one, send it to back, open the Selection Pane, check the model and a screenshot.
import { launch } from './ui_lib.mjs'
const t = await launch('arrange'); const { win } = t
await win.waitForFunction(() => globalThis.__ui)
await win.evaluate(() => globalThis.store.newWorkbook())
const menuClick = async (text) => { await win.locator('.xl-menu .xl-menu-item').filter({ hasText: text }).first().evaluate((e) => e.click()); await win.waitForTimeout(250) }
for (const [i, name] of ['Rectangle', 'Oval', 'Isosceles Triangle'].entries()) {
  await t.sel(2, 1 + i * 2)
  await t.tab('Insert')
  await win.locator('.xl-ribbon-groups button', { hasText: /^Shapes/ }).first().evaluate((e) => e.click()); await win.waitForTimeout(250)
  await menuClick(name)
}
const shapes = () => win.evaluate(() => (globalThis.store.active.layout.shapes ?? []).map((s) => ({ n: s.name, z: s.z, rot: s.rotation, flipH: s.flipH })))
t.log('inserted', await shapes())
await t.tab('Page Layout')
await win.locator('.xl-ribbon-groups button', { hasText: /^Rotate/ }).first().evaluate((e) => e.click()); await win.waitForTimeout(250)
await menuClick('Rotate Right 90')
t.log('after rotate', await shapes())
await win.locator('.xl-ribbon-groups button', { hasText: /^Send Backward/ }).first().evaluate((e) => e.click()); await win.waitForTimeout(250)
t.log('after send backward', await shapes())
await win.locator('.xl-ribbon-groups button', { hasText: /^Selection Pane/ }).first().evaluate((e) => e.click()); await win.waitForTimeout(400)
t.log('selection pane rows', await win.evaluate(() => document.querySelectorAll('.xl-selpane, .xl-selection-pane').length))
await t.shot('arrange')
t.log('errors', t.errors)
await t.close()

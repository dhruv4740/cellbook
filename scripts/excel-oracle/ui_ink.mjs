// Draw tab: pick the Red pen, draw a stroke, check the model and a screenshot.
import { launch } from './ui_lib.mjs'
const t = await launch('ink'); const { win } = t
await t.tab('Draw')
await win.locator('.xl-ribbon-groups button[aria-label*="Red"], .xl-ribbon-groups button[title*="Red"]').last().evaluate((e) => e.click()); await win.waitForTimeout(200)
t.log('tool', await win.evaluate(() => globalThis.__ui.getUi().inkTool))
await win.mouse.move(300, 400); await win.mouse.down()
for (let i = 0; i < 20; i++) await win.mouse.move(300 + i * 12, 400 + Math.sin(i / 2) * 40)
await win.mouse.up(); await win.waitForTimeout(300)
t.log('shapes', await win.evaluate(() => (globalThis.store.active.layout.shapes ?? []).map((s) => [s.kind, (s.points ?? []).length, s.stroke ?? s.color])))
await t.shot('ink'); t.log('errors', t.errors); await t.close()

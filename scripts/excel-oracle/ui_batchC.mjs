// Opens the not-yet-seen batch B controls and screenshots each; logs errors and visible dialog text.
import { launch } from './ui_lib.mjs'
const t = await launch('batchC'); const { win } = t
await win.evaluate(() => { const s = globalThis.store; s.newWorkbook(); const sh = s.active; ['a','b','c'].forEach((v,r)=>s.setRaw(sh,r,0,v+r)) })
const steps = [
  ['Page Layout', /^Themes$/, 'themes'], ['Insert', /^Text$/, 'textmenu'], ['Draw', /^Add Pen$|^Pens?$/, 'pens'],
  ['Home', /^Cell Styles$/, 'cellstyles'], ['Review', /^Allow Edit Ranges$|^Allow Edit/, 'editranges'], ['View', /^Custom Views$/, 'customviews'],
  ['Insert', /^Icons$/, 'icons'], ['Insert', /^Combo|^Recommended Charts$/, 'combo']
]
for (const [tab, re, name] of steps) {
  await t.tab(tab)
  const b = win.locator('button').filter({ hasText: re }).last()
  const n = await win.locator('button').filter({ hasText: re }).count()
  if (!n) { t.log(name, 'no button'); continue }
  await b.evaluate((e) => e.click()); await win.waitForTimeout(500)
  t.log(name, await win.evaluate(() => (document.querySelector('.xl-dialog')?.innerText ?? [...document.querySelectorAll('[role=menu],.xl-menu')].map((e) => e.innerText).join('|')).replace(/\s+/g, ' ').slice(0, 160)))
  await t.shot('c_' + name)
  await win.keyboard.press('Escape'); await win.waitForTimeout(200)
  await win.evaluate(() => document.querySelector('.xl-dialog-close,[aria-label=Close]')?.click())
}
t.log('errors', t.errors); await t.close()

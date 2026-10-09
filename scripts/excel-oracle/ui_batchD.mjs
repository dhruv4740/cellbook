// Drives the batch-B controls ui_batchC missed (label lookup by text / aria-label / title) and Split; screenshots each.
import { launch } from './ui_lib.mjs'
const t = await launch('batchD'); const { win } = t
await win.evaluate(() => { const s = globalThis.store; s.newWorkbook(); const sh = s.active; [['Item','Qty'],['a',3],['b',-2],['c',5],['d',4]].forEach((row,r)=>row.forEach((v,c)=>s.setRaw(sh,r,c,String(v)))); s.setSelection({ ranges: [{ r1:0,c1:0,r2:4,c2:1 }], active:{r:0,c:0} }) })
const click = (label) => win.evaluate((label) => {
  const m = (e) => (e.getAttribute('aria-label') || e.title || e.innerText || '').replace(/\s+/g, ' ').trim()
  const els = [...document.querySelectorAll('button')].filter((e) => { const x = m(e); return x === label || x.startsWith(label + ' (') })
  const e = els.at(-1); if (e) e.click(); return !!e
}, label)
const dlg = () => win.evaluate(() => (document.querySelector('.xl-dialog')?.innerText ?? '').replace(/\s+/g, ' ').slice(0, 140))
const menu = () => win.evaluate(() => [...document.querySelectorAll('[role=menuitem],[role=menuitemcheckbox],.xl-menu-item')].map((e) => e.innerText.replace(/\s+/g, ' ').trim()).slice(0, 14))
const close = async () => { await win.keyboard.press('Escape'); await win.waitForTimeout(150); await win.evaluate(() => [...document.querySelectorAll('.xl-dialog button')].find((b) => /^(Cancel|Close)$/.test(b.innerText.trim()))?.click()); await win.waitForTimeout(150) }
const step = async (tab, label, name, what = 'dlg') => {
  await t.tab(tab); await win.waitForTimeout(700)
  const ok = await click(label); await win.waitForTimeout(600)
  t.log(name, ok ? (what === 'dlg' ? await dlg() : await menu()) : 'NO BUTTON')
  await t.shot('d_' + name); await close()
}
await step('Draw', 'Add', 'addpen')
await step('Insert', 'Insert Combo Chart', 'combo', 'menu')
await step('Insert', 'Text', 'textmenu', 'menu')
await step(/^View/, 'Custom Views', 'customviews')
await step('Home', 'Format as Table', 'tablestyle', 'menu')
await step('Draw', 'Ink to Shape', 'inktoshape', 'menu')
// Split
await t.tab(/^View/); await t.sel(4, 2)
await click('Split'); await win.waitForTimeout(500)
t.log('split on', await win.evaluate(() => [globalThis.store.active.layout.frozen, globalThis.store.active.layout.split]))
await t.shot('d_split')
await click('Split'); await win.waitForTimeout(400)
t.log('split off', await win.evaluate(() => [globalThis.store.active.layout.frozen, globalThis.store.active.layout.split]))
await click('Split'); await click('Freeze Panes'); await win.waitForTimeout(300)
await win.evaluate(() => [...document.querySelectorAll('[role=menuitem],.xl-menu-item')].find((e) => /^Unfreeze/.test(e.innerText.trim()))?.click())
t.log('after unfreeze', await win.evaluate(() => [globalThis.store.active.layout.frozen, globalThis.store.active.layout.split]))
t.log('errors', t.errors); await t.close()

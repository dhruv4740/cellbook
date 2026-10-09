// Copy as Picture: captures the PNG handed to the clipboard IPC and saves it to out/ui/copypic.png.
import fs from 'fs'
import { launch } from './ui_lib.mjs'
const t = await launch('copypic'); const { win } = t
await win.evaluate(() => {
  const s = globalThis.store; s.newWorkbook(); const sh = s.active
  ;[['Item', 'Qty', 'Price'], ['apple', 3, 1.5], ['pear', 5, 2.25], ['plum', 2, 4]].forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v))))
  s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 3, c2: 2 }], active: { r: 0, c: 0 } })
  window.__png = null
  const orig = window.api.clipboardWriteImage
  window.api.clipboardWriteImage = async (u) => { window.__png = u; return orig(u) }
})
await t.tab('Home'); await win.waitForTimeout(700)
await win.evaluate(() => { const b = [...document.querySelectorAll('button')].filter((x) => /^Copy/.test(x.getAttribute('aria-label') || x.title || x.innerText)).at(-1); const arrow = b?.nextElementSibling ?? b?.parentElement?.querySelector('[class*=chev],[class*=arrow]'); (arrow ?? b)?.click() })
await win.waitForTimeout(700)
t.log('menu', JSON.stringify(await win.evaluate(() => [...document.querySelectorAll('[role=menuitem],.xl-menu-item')].map((e) => e.innerText.trim()))))
await win.evaluate(() => { [...document.querySelectorAll('[role=menuitem],.xl-menu-item')].find((e) => /Picture/i.test(e.innerText))?.click() })
await win.waitForTimeout(800)
const u = await win.evaluate(() => window.__png)
t.log('png', u ? u.length : 'none')
if (u) fs.writeFileSync('scripts/excel-oracle/out/ui/copypic.png', Buffer.from(u.split(',')[1], 'base64'))
t.log('errors', JSON.stringify(t.errors))
await t.close()

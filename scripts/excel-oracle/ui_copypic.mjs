// Copy as Picture: reads the PNG back from the OS clipboard and saves it to out/ui/copypic.png.
import { execSync } from 'child_process'
import { launch } from './ui_lib.mjs'
const t = await launch('copypic'); const { win } = t
await win.evaluate(() => {
  const s = globalThis.store; s.newWorkbook(); const sh = s.active
  ;[['Item', 'Qty', 'Price'], ['apple', 3, 1.5], ['pear', 5, 2.25], ['plum', 2, 4]].forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v))))
  s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 3, c2: 2 }], active: { r: 0, c: 0 } })
})
await t.tab('Home'); await win.waitForTimeout(700)
await win.evaluate(() => { const b = [...document.querySelectorAll('button')].filter((x) => /^Copy/.test(x.getAttribute('aria-label') || x.title || x.innerText)).at(-1); const arrow = b?.nextElementSibling ?? b?.parentElement?.querySelector('[class*=chev],[class*=arrow]'); (arrow ?? b)?.click() })
await win.waitForTimeout(700)
t.log('menu', JSON.stringify(await win.evaluate(() => [...document.querySelectorAll('[role=menuitem],.xl-menu-item')].map((e) => e.innerText.trim()))))
await win.evaluate(() => { [...document.querySelectorAll('[role=menuitem],.xl-menu-item')].find((e) => /Picture/i.test(e.innerText))?.click() })
await win.waitForTimeout(800)
import path from 'path'
const out = path.resolve('scripts/excel-oracle/out/ui/copypic.png')
const ps = `Add-Type -AssemblyName System.Windows.Forms; $i=[System.Windows.Forms.Clipboard]::GetImage(); if($i){$i.Save('${out}'); "$($i.Width)x$($i.Height)"} else {'no image'}`
t.log('clipboard image', execSync('powershell -NoProfile -EncodedCommand ' + Buffer.from(ps, 'utf16le').toString('base64')).toString().trim())
t.log('errors', JSON.stringify(t.errors))
await t.close()

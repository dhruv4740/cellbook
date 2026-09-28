import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-pu2-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.locator('.xl-tab', { hasText: 'Data' }).first().click()
await win.waitForTimeout(300)
await win.evaluate(() => {
  window.__ev = []
  for (const t of ['mousedown', 'mouseup', 'click', 'pointerdown', 'pointerup']) document.addEventListener(t, (e) => window.__ev.push(t + ':' + String(e.target.className || e.target.tagName).slice(0, 30)), true)
})
await win.getByText('Data Tools', { exact: true }).filter({ visible: true }).first().click()
await win.waitForTimeout(400)
await win.evaluate(() => { window.__ev.push('--- popup open: ' + !!document.querySelector('.xl-group-popup')) })
const box = await win.locator('.xl-group-popup').getByText('Consolidate').boundingBox()
console.log('label box', JSON.stringify(box))
await win.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
await win.waitForTimeout(200)
await win.evaluate(() => { window.__ev.push('--- after hover, popup open: ' + !!document.querySelector('.xl-group-popup')) })
await win.evaluate(() => { window.__dl = []; let last = 'x'; setInterval(() => { const d = JSON.stringify(globalThis.__ui.getUi().dialog); if (d !== last) { window.__dl.push(Math.round(performance.now()) + ':' + d); last = d } }, 1) })
await win.mouse.down()
await win.evaluate(() => { window.__ev.push('--- after down, popup open: ' + !!document.querySelector('.xl-group-popup')) })
await win.mouse.up()
await win.waitForTimeout(400)
console.log((await win.evaluate(() => window.__ev)).join('\n'))
console.log('dialog history', JSON.stringify(await win.evaluate(() => window.__dl)))
console.log('ui.dialog:', JSON.stringify(await win.evaluate(() => globalThis.__ui.getUi().dialog)), 'popup open:', await win.evaluate(() => !!document.querySelector('.xl-group-popup')))
console.log('dialog:', await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent ?? 'none'))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

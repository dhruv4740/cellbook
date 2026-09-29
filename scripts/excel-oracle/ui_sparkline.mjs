// Insert > Sparklines through the real dialog, then a screenshot of the grid.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out'); fs.mkdirSync(out, { recursive: true })
const profile = path.resolve(out, 'rt/profile-sp-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.evaluate(() => { const s = globalThis.store; const sh = s.active; for (let r = 0; r < 4; r++) for (let c = 0; c < 6; c++) s.setRaw(sh, r, c, String(((r + 2) * (c + 3)) % 7 - 2)) })
await win.locator('.xl-tab', { hasText: /^Insert/ }).first().click()
await win.locator('button', { hasText: /^Line$/ }).filter({ visible: true }).first().click()
await win.waitForTimeout(400)
const inputs = win.locator('.xl-dialog input')
await inputs.nth(0).fill('A1:F4'); await inputs.nth(1).fill('H1:H4')
await win.keyboard.press('Enter')
await win.waitForTimeout(500)
console.log('sparklines', await win.evaluate(() => globalThis.store.active.layout.sparklines?.length))
await win.evaluate(() => { globalThis.__shot = true })
const b64 = await app.evaluate(async ({ BrowserWindow }) => {
  let t = BrowserWindow.getAllWindows()[0]
  for (const w of BrowserWindow.getAllWindows()) if (await w.webContents.executeJavaScript('!!globalThis.__shot').catch(() => false)) t = w
  return (await t.webContents.capturePage()).toPNG().toString('base64')
})
fs.writeFileSync(path.join(out, 'sparkline.png'), Buffer.from(b64, 'base64'))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

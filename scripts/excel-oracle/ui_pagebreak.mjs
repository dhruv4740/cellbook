// View > Page Break Preview: screenshot of a sheet spanning several pages.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out'); fs.mkdirSync(out, { recursive: true })
const profile = path.resolve(out, 'rt/profile-pb-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store && globalThis.__ui)
await win.evaluate(() => { const s = globalThis.store; for (let r = 0; r < 120; r++) for (let c = 0; c < 14; c++) s.setRaw(s.active, r, c, 'r' + r + 'c' + c) })
await win.locator('.xl-tab', { hasText: /^View/ }).first().click()
await win.locator('button', { hasText: /^Page Break Preview/ }).filter({ visible: true }).first().click()
await win.waitForTimeout(800)
console.log('mode', await win.evaluate(() => globalThis.__ui.getUi().viewMode))
await win.evaluate(() => { globalThis.__shot = true })
const b64 = await app.evaluate(async ({ BrowserWindow }) => {
  let t = BrowserWindow.getAllWindows()[0]
  for (const w of BrowserWindow.getAllWindows()) if (await w.webContents.executeJavaScript('!!globalThis.__shot').catch(() => false)) t = w
  return (await t.webContents.capturePage()).toPNG().toString('base64')
})
fs.writeFileSync(path.join(out, 'pagebreak.png'), Buffer.from(b64, 'base64'))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

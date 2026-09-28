// One screenshot of Cellbook's ribbon per tab (compare with ribbon_shots.py, which captures real Excel).
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out/ribbon'); fs.mkdirSync(out, { recursive: true })
const profile = path.resolve('scripts/excel-oracle/out/rt/profile-rb-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
const size = await win.evaluate(() => [window.innerWidth, window.innerHeight])
console.log('window', size)
// (no reduced-motion emulation: it re-lays out the ribbon)
const shot = async (n) => {
  const b64 = await app.evaluate(async ({ BrowserWindow }) => {
    let target = BrowserWindow.getAllWindows()[0]
    for (const w of BrowserWindow.getAllWindows()) { const ok = await w.webContents.executeJavaScript('!!globalThis.__ribbonShot').catch(() => false); if (ok) target = w }
    return (await target.webContents.capturePage()).toPNG().toString('base64')
  })
  fs.writeFileSync(path.join(out, 'cellbook-' + n + '.png'), Buffer.from(b64, 'base64'))
}
await win.evaluate(() => { globalThis.__ribbonShot = true })
for (const [name, label] of [['home', 'Home'], ['insert', 'Insert'], ['draw', 'Draw'], ['pagelayout', 'Page Layout'], ['formulas', 'Formulas'], ['data', 'Data'], ['review', 'Review'], ['view', 'View'], ['help', 'Help']]) {
  await win.locator('.xl-tab', { hasText: new RegExp('^' + label) }).first().click()
  await win.mouse.move(600, 600)
  await win.waitForTimeout(1200)
  await shot(name)
  console.log('saved', name)
}
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

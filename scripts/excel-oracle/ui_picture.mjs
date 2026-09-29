// Insert > Pictures > This Device through the real UI (file chooser), then a screenshot of the grid.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out'); fs.mkdirSync(out, { recursive: true })
const profile = path.resolve(out, 'rt/profile-pic-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
const dataUrl = await win.evaluate(() => {
  const c = document.createElement('canvas'); c.width = 240; c.height = 120
  const g = c.getContext('2d'); const grad = g.createLinearGradient(0, 0, 240, 120); grad.addColorStop(0, '#217346'); grad.addColorStop(1, '#f2b134')
  g.fillStyle = grad; g.fillRect(0, 0, 240, 120); g.fillStyle = '#fff'; g.font = '28px sans-serif'; g.fillText('LOGO', 80, 70)
  return c.toDataURL('image/png')
})
const pngPath = path.join(out, 'test-logo.png')
fs.writeFileSync(pngPath, Buffer.from(dataUrl.split(',')[1], 'base64'))
await win.evaluate(() => { const s = globalThis.store; s.setRaw(s.active, 0, 0, 'Report'); s.setSelection({ ranges: [{ r1: 2, c1: 2, r2: 2, c2: 2 }], active: { r: 2, c: 2 } }) })
await win.locator('.xl-tab', { hasText: /^Insert/ }).first().click()
await win.locator('button', { hasText: /^Pictures/ }).filter({ visible: true }).first().click()
const chooser = win.waitForEvent('filechooser')
await win.getByText('This Device...').click()
;(await chooser).setFiles(pngPath)
await win.waitForTimeout(800)
console.log('pictures', await win.evaluate(() => JSON.stringify(globalThis.store.active.layout.pictures?.map((p) => ({ n: p.name, a: p.anchor, w: p.width, h: p.height })))))
// drag it 100px right/down and check the anchor moved
const box = await win.locator('img[alt^="Picture"]').first().boundingBox()
await win.mouse.move(box.x + 60, box.y + 40); await win.mouse.down(); await win.mouse.move(box.x + 160, box.y + 90, { steps: 5 }); await win.mouse.up()
await win.waitForTimeout(300)
console.log('after drag', await win.evaluate(() => JSON.stringify(globalThis.store.active.layout.pictures?.[0].anchor)))
await win.keyboard.press('Control+z')
await win.waitForTimeout(200)
console.log('after undo', await win.evaluate(() => JSON.stringify(globalThis.store.active.layout.pictures?.[0].anchor)))
await win.evaluate(() => { globalThis.__shot = true })
const b64 = await app.evaluate(async ({ BrowserWindow }) => {
  let t = BrowserWindow.getAllWindows()[0]
  for (const w of BrowserWindow.getAllWindows()) if (await w.webContents.executeJavaScript('!!globalThis.__shot').catch(() => false)) t = w
  return (await t.webContents.capturePage()).toPNG().toString('base64')
})
fs.writeFileSync(path.join(out, 'picture.png'), Buffer.from(b64, 'base64'))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

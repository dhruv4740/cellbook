// Insert > Shapes / Text Box through the real ribbon, edit text by double-click, screenshot.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out'); fs.mkdirSync(out, { recursive: true })
const profile = path.resolve(out, 'rt/profile-shp-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
const errors = []
win.on('pageerror', (e) => errors.push(String(e)))
await win.locator('.xl-tab', { hasText: /^Insert/ }).first().click()
const pick = async (btn, item) => {
  await win.locator('button', { hasText: btn }).filter({ visible: true }).first().click()
  if (item) await win.getByText(item, { exact: true }).click()
  await win.waitForTimeout(300)
}
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 2, c1: 1, r2: 2, c2: 1 }], active: { r: 2, c: 1 } }))
await pick(/^Shapes/, 'Rectangle: Rounded Corners')
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 2, c1: 5, r2: 2, c2: 5 }], active: { r: 2, c: 5 } }))
await pick(/^Shapes/, 'Arrow: Right')
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 9, c1: 1, r2: 9, c2: 1 }], active: { r: 9, c: 1 } }))
await pick(/^Text$/, 'Text Box')
console.log('shapes', await win.evaluate(() => JSON.stringify(globalThis.store.active.layout.shapes?.map((s) => s.kind + ':' + s.name))))
// type into the text box (selected after insert): double-click, type, click away
const tb = win.locator('svg').last()
const box = await win.locator('.grid-host svg').last().boundingBox().catch(() => null)
await win.mouse.dblclick(box ? box.x + 30 : 200, box ? box.y + 20 : 500)
await win.waitForTimeout(300)
await win.keyboard.type('Dispatch notes')
await win.mouse.click(700, 700)
await win.waitForTimeout(300)
console.log('text', await win.evaluate(() => JSON.stringify(globalThis.store.active.layout.shapes?.map((s) => s.text))))
await win.evaluate(() => { globalThis.__shot = true })
const b64 = await app.evaluate(async ({ BrowserWindow }) => {
  let t = BrowserWindow.getAllWindows()[0]
  for (const w of BrowserWindow.getAllWindows()) if (await w.webContents.executeJavaScript('!!globalThis.__shot').catch(() => false)) t = w
  return (await t.webContents.capturePage()).toPNG().toString('base64')
})
fs.writeFileSync(path.join(out, 'shapes.png'), Buffer.from(b64, 'base64'))
console.log('errors:', errors.length ? errors.join(' || ').slice(0, 500) : 'none')
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

// Drives the built app with real mouse/keyboard for the outline + protection UI. node scripts/excel-oracle/ui_check.mjs
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const outDir = path.resolve('scripts/excel-oracle/out/ui'); fs.mkdirSync(outDir, { recursive: true })
const profile = path.resolve('scripts/excel-oracle/out/rt/profile-ui-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
const log = (k, v) => console.log(k, JSON.stringify(v))
const shot = async (n) => {
  const b64 = await app.evaluate(async ({ BrowserWindow }) => { const w = BrowserWindow.getAllWindows().find((x) => x.isVisible()) ?? BrowserWindow.getAllWindows()[0]; return (await w.webContents.capturePage()).toPNG().toString('base64') })
  fs.writeFileSync(path.join(outDir, n + '.png'), Buffer.from(b64, 'base64'))
}
await win.evaluate(() => { for (let r = 0; r < 12; r++) { store.setRaw(store.active, r, 0, String(r + 1)); store.setRaw(store.active, r, 1, String((r + 1) * 10)) } })
// --- outline via keyboard: select rows 2:4 (shift+space) and Alt+Shift+Right
await win.evaluate(() => store.setSelection({ ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], active: { r: 1, c: 0 } }))
await win.locator('.grid-host').first().focus()
log('scrollLeft before', await win.evaluate(() => document.querySelector('.grid-host').scrollLeft ?? -1))
await win.keyboard.press('Shift+Space')
log('scrollLeft after Shift+Space', await win.evaluate(() => [...document.querySelectorAll('*')].filter((e) => e.scrollLeft > 0).map((e) => e.className + ':' + e.scrollLeft)))
await win.keyboard.press('Alt+Shift+ArrowRight')
await win.waitForTimeout(300)
log('levels after Alt+Shift+Right', await win.evaluate(() => store.active.layout.rowLevels))
await shot('outline-1-grouped')
// click the [-] button: gutter geometry -> summary row 5 (index 4); find the canvas position
const host = await win.locator('.grid-host').first().boundingBox()
// header band 20px + gutter; rows are 20px high; summary row index 4 center
const x = host.x + 2 + 5.5
await win.mouse.click(x, host.y + 20 + 4 * 20 + 10)
await win.waitForTimeout(300)
log('hidden after clicking [-]', await win.evaluate(() => Object.keys(store.active.layout.hiddenRows)))
await shot('outline-2-collapsed')
await win.mouse.click(x, host.y + 20 + 1 * 20 + 10) // summary row moved up once rows 2:4 are hidden
await win.waitForTimeout(200)
log('hidden after clicking [+]', await win.evaluate(() => Object.keys(store.active.layout.hiddenRows)))
// level button "1" (top-left corner)
await win.mouse.click(host.x + 2 + 6, host.y + 10)
await win.waitForTimeout(200)
log('hidden after level button 1', await win.evaluate(() => Object.keys(store.active.layout.hiddenRows)))
// --- protect via ribbon
await win.locator('.xl-tab', { hasText: 'Review' }).first().click()
await win.waitForTimeout(300)
await shot('review-tab')
await win.getByText('Protect Sheet').filter({ visible: true }).first().click()
await win.waitForTimeout(300)
await shot('protect-1-dialog')
await win.keyboard.press('Enter')
await win.waitForTimeout(300)
log('protected', await win.evaluate(() => !!store.active.layout.protection))
await win.locator('.grid-host').first().focus()
await win.evaluate(() => store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }))
await win.keyboard.type('hello')
await win.waitForTimeout(300)
await shot('protect-2-message')
log('A1 after typing on protected sheet', await win.evaluate(() => store.getRaw(store.active, 0, 0)))
log('dialog', await win.evaluate(() => document.querySelector('.xl-dialog')?.textContent))
await win.keyboard.press('Enter')
await app.evaluate(({ app }) => app.exit(0)).catch(() => {})
process.exit(0)

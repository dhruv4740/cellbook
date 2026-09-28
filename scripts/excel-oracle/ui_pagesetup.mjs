// Drives the built app: Page Layout tab menus, Page Setup dialog, print preview. node scripts/excel-oracle/ui_pagesetup.mjs
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const outDir = path.resolve('scripts/excel-oracle/out/ui'); fs.mkdirSync(outDir, { recursive: true })
const profile = path.resolve('scripts/excel-oracle/out/rt/profile-ps-' + Date.now())
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
await win.evaluate(() => { for (let r = 0; r < 80; r++) for (let c = 0; c < 4; c++) store.setRaw(store.active, r, c, 'r' + r + 'c' + c) })
await win.locator('.xl-tab', { hasText: 'Page Layout' }).first().click()
await win.waitForTimeout(300)
await win.getByText('Margins').filter({ visible: true }).first().click()
await win.waitForTimeout(200)
await shot('ps-1-margins-menu')
await win.getByText('Narrow').filter({ visible: true }).first().click()
await win.waitForTimeout(200)
log('margins', await win.evaluate(() => store.active.layout.pageSetup?.margins))
// launcher -> dialog
await win.evaluate(() => { const b = [...document.querySelectorAll('.xl-launcher, [class*=launcher]')].filter((e) => e.offsetParent); b[0]?.click() })
await win.waitForTimeout(300)
log('dialog title', await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent))
await win.getByText('Header/Footer', { exact: true }).click()
await win.waitForTimeout(200)
await shot('ps-2-headerfooter')
await win.locator('.xl-dialog select').first().selectOption({ label: 'Page 1 of ?' })
await win.getByText('Sheet', { exact: true }).last().click()
await win.locator('.xl-dialog input').nth(1).fill('$1:$1')
await win.keyboard.press('Enter')
await win.waitForTimeout(300)
log('hf', await win.evaluate(() => store.active.layout.pageSetup?.hf))
log('titles', await win.evaluate(() => store.printTitles(store.active)))
// page break + preview
await win.evaluate(() => store.setSelection({ ranges: [{ r1: 20, c1: 1, r2: 20, c2: 1 }], active: { r: 20, c: 1 } }))
await win.getByText('Breaks').filter({ visible: true }).first().click()
await win.getByText('Insert Page Break').filter({ visible: true }).first().click()
await win.waitForTimeout(200)
log('breaks', await win.evaluate(() => [store.active.layout.pageSetup?.rowBreaks, store.active.layout.pageSetup?.colBreaks]))
await win.evaluate(() => globalThis.setUi?.({ backstage: true, backstageTab: 'print' }))
await win.keyboard.press('Control+P')
await win.waitForTimeout(800)
await shot('ps-3-print-preview')
log('preview iframe', await win.evaluate(() => document.querySelector('.bs-print-frame')?.srcdoc.length))
await app.close()

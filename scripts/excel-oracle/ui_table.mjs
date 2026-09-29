// Opens the Excel-made table workbook (table_rt.py writes out/table_src.xlsx) and screenshots the grid: the table style should show.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out')
const file = path.join(out, 'table_src.xlsx')
const profile = path.resolve(out, 'rt/profile-tbl-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.evaluate((p) => window.api.openPathRequest(p), file)
let data
for (let i = 0; i < 100 && !data; i++) {
  for (const p of app.windows()) if (await p.evaluate(() => globalThis.store?.getRaw?.(globalThis.store.active, 0, 0) === 'Item').catch(() => false)) data = p
  if (!data) await new Promise((r) => setTimeout(r, 200))
}
console.log('opened', !!data, 'tables', await data.evaluate(() => JSON.stringify(globalThis.store.active.layout.tables)))
await data.evaluate(() => { globalThis.__shot = true })
await data.waitForTimeout(500)
const b64 = await app.evaluate(async ({ BrowserWindow }) => {
  let t = BrowserWindow.getAllWindows()[0]
  for (const w of BrowserWindow.getAllWindows()) if (await w.webContents.executeJavaScript('!!globalThis.__shot').catch(() => false)) t = w
  return (await t.webContents.capturePage()).toPNG().toString('base64')
})
fs.writeFileSync(path.join(out, 'table.png'), Buffer.from(b64, 'base64'))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

// Review > Spelling (F7) through the real dialog: Change, Ignore All, finish.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out'); fs.mkdirSync(out, { recursive: true })
const profile = path.resolve(out, 'rt/profile-sp2-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.evaluate(() => {
  const s = globalThis.store; const sh = s.active
  s.setRaw(sh, 0, 0, 'Shiping adress'); s.setRaw(sh, 1, 0, 'all good here'); s.setRaw(sh, 2, 0, 'Xyzzyq station')
  s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } })
})
await win.locator('.grid-host').first().focus()
await win.keyboard.press('F7')
await win.waitForSelector('.xl-dialog-title')
const title = () => win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent)
const shown = () => win.evaluate(() => [...document.querySelectorAll('.xl-dialog span')].map((s) => s.textContent).join('|'))
await win.waitForFunction(() => document.querySelector('.xl-dialog span')?.textContent)
console.log('1', await title(), await shown())
await win.getByRole('button', { name: 'Change', exact: true }).click()
await win.waitForTimeout(300)
console.log('2', await shown(), '| A1 =', await win.evaluate(() => globalThis.store.getRaw(globalThis.store.active, 0, 0)))
await win.getByRole('button', { name: 'Ignore All' }).click()
await win.waitForTimeout(300)
console.log('3', await shown())
await win.getByRole('button', { name: 'Ignore Once' }).click()
await win.waitForTimeout(300)
console.log('4', await title(), '|', await win.evaluate(() => document.querySelector('.xl-dialog')?.textContent))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

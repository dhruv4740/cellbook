// Opens the AutoFilter dropdown on a big sheet (keyboard: Alt+Down on a header cell) and reports timing and console errors.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const file = path.resolve(process.argv[2] ?? 'samples/private/scaled-50000.xlsx')
const col = Number(process.argv[3] ?? 4)
const profile = path.resolve('scripts/out/profile-fm-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.evaluate((p) => window.api.openPathRequest(p), file)
let data
for (let i = 0; i < 600 && !data; i++) { for (const p of app.windows()) { const rows = await p.evaluate(() => { const u = globalThis.store?.usedRange?.(globalThis.store.active); return u ? u.r2 + 1 : 0 }).catch(() => 0); if (rows > 1000) data = p } if (!data) await new Promise((r) => setTimeout(r, 100)) }
win = data
win.on('console', (m) => { if (m.type() === 'error' || m.type() === 'warning') console.log('console', m.type(), m.text().slice(0, 300)) })
win.on('pageerror', (e) => console.log('pageerror', String(e).slice(0, 300)))
await win.waitForTimeout(1000)
await win.evaluate(() => { const s = globalThis.store; s.setSelection({ ranges: [{ r1: 2, c1: 3, r2: 2, c2: 3 }], active: { r: 2, c: 3 } }); s.toggleAutoFilter(s.active) })
console.log('autofilter', JSON.stringify(await win.evaluate(() => globalThis.store.active.layout.autoFilter?.range)))
await win.evaluate((c) => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: c, r2: 0, c2: c }], active: { r: 0, c } }), col)
await win.locator('.grid-host').first().focus()
const t = Date.now()
await win.keyboard.press('Alt+ArrowDown')
const ok = await win.locator('.xl-afmenu').first().waitFor({ timeout: 60000 }).then(() => true).catch(() => false)
console.log('menu visible:', ok, 'after', Date.now() - t, 'ms; items', await win.locator('.xl-afmenu-check').count())
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

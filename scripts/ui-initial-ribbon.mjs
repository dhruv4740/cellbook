import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-ir-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
const read = () => win.evaluate(() => { const host = document.querySelector('.xl-ribbon-groups'); return { avail: host.clientWidth, groups: [...host.children].filter((c) => c.getBoundingClientRect().width > 2).map((c) => Math.round(c.getBoundingClientRect().width) + ':' + (c.textContent || '').slice(0, 10)) } })
for (const ms of [200, 1000, 3000]) { await win.waitForTimeout(ms); console.log('t+', ms, JSON.stringify(await read())) }
await win.locator('.xl-tab', { hasText: /^Insert/ }).first().click()
await win.locator('.xl-tab', { hasText: /^Home/ }).first().click()
await win.waitForTimeout(500)
console.log('after tab switch', JSON.stringify(await read()))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

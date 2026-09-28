import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-an-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.locator('.xl-tab', { hasText: 'Data' }).first().click()
for (const ms of [50, 200, 600, 1500]) {
  await win.waitForTimeout(ms === 50 ? 50 : ms - (ms === 200 ? 50 : ms === 600 ? 200 : 600))
  console.log(ms, JSON.stringify(await win.evaluate(() => { const g = document.querySelector('.xl-ribbon-groups'); return { opacity: getComputedStyle(g).opacity, anims: document.getAnimations().map((a) => a.playState + ':' + Math.round(a.currentTime ?? -1)), vis: document.visibilityState, focus: document.hasFocus() } })))
}
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

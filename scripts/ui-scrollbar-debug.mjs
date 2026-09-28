import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-sb-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.waitForTimeout(800)
console.log(JSON.stringify(await win.evaluate(() => {
  const s = document.querySelector('.xl-scroll'), c = document.querySelector('.xl-canvas'), h = document.querySelector('.grid-host')
  const cs = getComputedStyle(s)
  return { scroll: { off: [s.offsetWidth, s.offsetHeight], client: [s.clientWidth, s.clientHeight], sw: [s.scrollWidth, s.scrollHeight], overflow: cs.overflow, sbw: cs.scrollbarWidth }, canvas: [c.clientWidth, c.clientHeight, getComputedStyle(c).zIndex], host: [h.clientWidth, h.clientHeight], dpr: devicePixelRatio }
})))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

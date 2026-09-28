import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const flags = process.argv[2] === 'flags' ? ['--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] : []
const profile = path.resolve('scripts/out/profile-raf-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, ...flags] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
console.log('vis', await win.evaluate(() => document.visibilityState), 'focus', await win.evaluate(() => document.hasFocus()))
const r = await Promise.race([win.evaluate(() => new Promise((res) => requestAnimationFrame(() => res('raf ok')))), new Promise((res) => setTimeout(() => res('raf TIMEOUT'), 5000))])
console.log(flags.length ? 'with flags:' : 'no flags:', r)
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-sp-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
console.log('languages', await win.evaluate(() => window.api.spell.languages()))
for (const w of ['hello', 'teh', 'recieve', 'Excel', 'xyzzyq']) {
  console.log(w, JSON.stringify(await win.evaluate((w) => ({ bad: window.api.spell.isMisspelled(w), sug: window.api.spell.suggestions(w).slice(0, 4) }), w)))
}
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

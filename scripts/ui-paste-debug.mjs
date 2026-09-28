import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-pd-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
await win.evaluate(() => { const s = globalThis.store; for (let r = 0; r < 200; r++) for (let c = 0; c < 5; c++) s.setRaw(s.active, r, c, `r${r}c${c}`) })
const active = async () => win.evaluate(() => document.activeElement?.className + ' | ' + document.activeElement?.tagName)
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 99, c2: 4 }], active: { r: 0, c: 0 } }))
await win.locator('.grid-host').first().focus()
console.log('active before copy:', await active())
await win.keyboard.press('Control+c')
await win.waitForTimeout(500)
console.log('clipboard set:', await win.evaluate(() => !!globalThis.store.clipboard), '| active after copy:', await active())
const sys = await win.evaluate(async () => { const t = await window.api.clipboardReadText(); return { type: typeof t, len: t?.length, head: String(t).slice(0, 40), stored: globalThis.store.clipboard?.text?.length } })
console.log('system clipboard after copy:', JSON.stringify(sys))
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 8, r2: 0, c2: 8 }], active: { r: 0, c: 8 } }))
console.log('active before paste:', await active())
let t = Date.now()
await win.keyboard.press('Control+v')
for (let i = 0; i < 100; i++) { if (await win.evaluate(() => globalThis.store.getRaw(globalThis.store.active, 99, 12) !== '')) break; await win.waitForTimeout(50) }
console.log('paste done:', await win.evaluate(() => globalThis.store.getRaw(globalThis.store.active, 99, 12)), (Date.now() - t) + ' ms')
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

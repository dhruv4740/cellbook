// CPU-profiles the renderer while doing one interaction on a big workbook.
//   node scripts/ui-profile.mjs samples/private/scaled-50000.xlsx arrows|scroll|type
// Writes scripts/out/ui-<mode>.cpuprofile (summarise with python scripts/prof.py after copying it as the newest file).
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'

const file = path.resolve(process.argv[2] ?? 'samples/private/scaled-50000.xlsx')
const mode = process.argv[3] ?? 'arrows'
const profile = path.resolve('scripts/out/profile-prof-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) {
  for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p
  if (!win) await new Promise((r) => setTimeout(r, 200))
}
await win.waitForFunction(() => globalThis.store)
await win.evaluate((p) => window.api.openPathRequest(p), file)
let data
for (let i = 0; i < 600 && !data; i++) {
  for (const p of app.windows()) {
    const rows = await p.evaluate(() => { const u = globalThis.store?.usedRange?.(globalThis.store.active); return u ? u.r2 + 1 : 0 }).catch(() => 0)
    if (rows > 1000) data = p
  }
  if (!data) await new Promise((r) => setTimeout(r, 100))
}
win = data
await win.waitForTimeout(1500)
const host = await win.locator('.grid-host').first().boundingBox()
await win.locator('.grid-host').first().click({ position: { x: 200, y: 120 } })
await win.waitForTimeout(500)
if (mode === 'paste' || mode === 'copy') {
  await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 7999, c2: 20 }], active: { r: 0, c: 0 } }))
  await win.locator('.grid-host').first().focus()
  if (mode === 'paste') { await win.keyboard.press('Control+c'); await win.waitForTimeout(3000) }
  await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 26, r2: 0, c2: 26 }], active: { r: 0, c: 26 } }))
  await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 7999, c2: 20 }], active: { r: 0, c: 0 } }))
}
const cdp = await app.context().newCDPSession(win)
await cdp.send('Profiler.enable')
await cdp.send('Profiler.setSamplingInterval', { interval: 200 })
await cdp.send('Profiler.start')
const t0 = Date.now()
if (mode === 'arrows') for (let i = 0; i < 12; i++) await win.keyboard.press('ArrowDown')
else if (mode === 'scroll') { await win.mouse.move(host.x + 300, host.y + 200); for (let i = 0; i < 40; i++) { await win.mouse.wheel(0, 240); await win.waitForTimeout(8) } }
else if (mode === 'copy') { await win.keyboard.press('Control+c'); await win.waitForTimeout(500) }
else if (mode === 'paste') {
  await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 26, r2: 0, c2: 26 }], active: { r: 0, c: 26 } }))
  await win.keyboard.press('Control+v')
  for (let i = 0; i < 1200; i++) { const done = await win.evaluate(() => globalThis.store.getRaw(globalThis.store.active, 7999, 26) !== ''); if (done) break; await win.waitForTimeout(50) }
}
else if (mode === 'type') { await win.keyboard.type('12345'); await win.keyboard.press('Enter') }
const ms = Date.now() - t0
const { profile: prof } = await cdp.send('Profiler.stop')
fs.mkdirSync('scripts/out', { recursive: true })
const out = `scripts/out/ui-${mode}.cpuprofile`
fs.writeFileSync(out, JSON.stringify(prof))
console.log(mode, 'took', ms, 'ms; wrote', out)
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

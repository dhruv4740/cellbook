// Drives the built app on a big workbook and measures what a user feels: open time, scrolling frame times, keyboard
// navigation, typing, AutoFilter and sort. Usage: node scripts/ui-perf.mjs samples/private/scaled-50000.xlsx
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'

const file = path.resolve(process.argv[2] ?? 'samples/private/scaled-50000.xlsx')
const profile = path.resolve('scripts/out/profile-perf-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) {
  for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p
  if (!win) await new Promise((r) => setTimeout(r, 200))
}
await win.waitForFunction(() => globalThis.store)
const log = (k, v) => console.log(k.padEnd(34), typeof v === 'string' ? v : JSON.stringify(v))

// ---- open ----
let t = Date.now()
await win.evaluate((p) => window.api.openPathRequest(p), file)
// the file may open in a new window; find the window that holds the data
let data
for (let i = 0; i < 600 && !data; i++) {
  for (const p of app.windows()) {
    const rows = await p.evaluate(() => { const u = globalThis.store?.usedRange?.(globalThis.store.active); return u ? u.r2 + 1 : 0 }).catch(() => 0)
    if (rows > 1000) data = p
  }
  if (!data) await new Promise((r) => setTimeout(r, 100))
}
if (!data) { console.log('file did not open'); await app.close(); process.exit(1) }
win = data
log('open (until data visible)', (Date.now() - t) + ' ms')
await win.waitForTimeout(500)
log('open timing (read in worker / store.load)', await win.evaluate(() => globalThis.__openTiming))
log('used range rows', await win.evaluate(() => globalThis.store.usedRange(globalThis.store.active).r2 + 1))

// ---- frame-time recorder ----
await win.evaluate(() => {
  window.__frames = []
  let last = performance.now()
  const tick = (now) => { window.__frames.push(now - last); last = now; requestAnimationFrame(tick) }
  requestAnimationFrame(tick)
})
const frameStats = async (label) => {
  const f = await win.evaluate(() => { const a = window.__frames.splice(0); return a })
  if (!f.length) return log(label, 'no frames')
  const s = [...f].sort((a, b) => a - b)
  const p = (q) => s[Math.min(s.length - 1, Math.floor(q * s.length))].toFixed(1)
  log(label, `frames ${f.length}  avg ${(f.reduce((a, b) => a + b, 0) / f.length).toFixed(1)}ms  p95 ${p(0.95)}  max ${s[s.length - 1].toFixed(1)}  >50ms: ${f.filter((x) => x > 50).length}`)
}
const host = await win.locator('.grid-host').first().boundingBox()
await win.mouse.move(host.x + 300, host.y + 200)
await frameStats('idle (discard)')
await win.waitForTimeout(300)
await frameStats('idle')

// ---- wheel scrolling ----
for (let i = 0; i < 80; i++) { await win.mouse.wheel(0, 240); await win.waitForTimeout(8) }
await win.waitForTimeout(200)
await frameStats('wheel scroll down x80')
for (let i = 0; i < 40; i++) { await win.mouse.wheel(240, 0); await win.waitForTimeout(8) }
await win.waitForTimeout(200)
await frameStats('wheel scroll right x40')
for (let i = 0; i < 80; i++) { await win.mouse.wheel(0, -480); await win.waitForTimeout(8) }
await win.waitForTimeout(200)
await frameStats('fast scroll up x80')

// ---- keyboard ----
await win.locator('.grid-host').first().click({ position: { x: 200, y: 120 } })
await frameStats('click discard')
t = Date.now()
for (let i = 0; i < 100; i++) await win.keyboard.press('ArrowDown')
log('100 x ArrowDown', (Date.now() - t) + ' ms')
t = Date.now()
for (let i = 0; i < 20; i++) await win.keyboard.press('PageDown')
log('20 x PageDown', (Date.now() - t) + ' ms')
t = Date.now()
await win.keyboard.press('Control+ArrowDown')
await win.waitForTimeout(50)
log('Ctrl+Down (to last row)', (Date.now() - t) + ' ms, active row ' + (await win.evaluate(() => globalThis.store.active.selection.active.r)))
t = Date.now()
await win.keyboard.press('Control+Home')
await win.waitForTimeout(50)
log('Ctrl+Home', (Date.now() - t) + ' ms')
await frameStats('keyboard (discard)')

// ---- typing / editing ----
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 2, c1: 9, r2: 2, c2: 9 }], active: { r: 2, c: 9 } }))
t = Date.now()
await win.keyboard.type('123')
await win.keyboard.press('Enter')
await win.waitForTimeout(20)
log('type 3 chars + Enter (edit a key col)', (Date.now() - t) + ' ms')
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 2, c1: 3, r2: 2, c2: 3 }], active: { r: 2, c: 3 } }))
t = Date.now()
await win.keyboard.type('CHANGED')
await win.keyboard.press('Enter')
await win.waitForTimeout(20)
log('edit a lookup-key cell', (Date.now() - t) + ' ms')
await win.keyboard.press('Control+z')
await win.waitForTimeout(20)

// ---- data operations (in the renderer, same engine the buttons call) ----
const time = async (label, fn) => {
  const ms = await win.evaluate(async (src) => { const f = new Function('store', 'return (' + src + ')(store)'); const t0 = performance.now(); await f(globalThis.store); return performance.now() - t0 }, fn.toString())
  log(label, ms.toFixed(0) + ' ms')
}
await time('turn on AutoFilter', (s) => s.toggleAutoFilter(s.active))
// open the AutoFilter dropdown of a header cell with Alt+Down and time until it is on screen
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 4, r2: 0, c2: 4 }], active: { r: 0, c: 4 } }))
await win.locator('.grid-host').first().focus()
t = Date.now()
await win.keyboard.press('Alt+ArrowDown')
await win.locator('.xl-afmenu').first().waitFor({ timeout: 20000 }).catch(() => {})
log('open filter dropdown (Alt+Down)', (Date.now() - t) + ' ms, list items: ' + (await win.locator('.xl-afmenu-check').count()))
await win.keyboard.press('Escape')
await win.locator('.xl-afmenu-scrim').first().click({ timeout: 1000 }).catch(() => {})
await time('sort via filter menu (asc, SOURCE)', (s) => s.sortFilterRange(s.active, 4, 'asc'))
await time('undo sort', (s) => s.undo())
await time('filter SOURCE = one value', (s) => { const v = s.getDisplay(s.active, 1, 4).text; s.setColumnFilter(s.active, 4, { selected: [v] }) })
await time('clear filters', (s) => s.clearAllFilters(s.active))
await time('turn off AutoFilter', (s) => s.toggleAutoFilter(s.active))
await time('insert 5 rows at row 10', (s) => s.insertRows(s.active, 10, 5))
await time('undo', (s) => s.undo())
await time('delete 5 rows at row 10', (s) => s.deleteRows(s.active, 10, 5))
await time('undo', (s) => s.undo())
await time('Select-all + bold (applyStyle)', (s) => s.applyStyle(s.active, [{ r1: 0, c1: 0, r2: s.usedRange(s.active).r2, c2: 20 }], { bold: true }))
await time('undo', (s) => s.undo())
// copy / paste 20,000 rows x 21 columns through the keyboard
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 19999, c2: 20 }], active: { r: 0, c: 0 } }))
await win.locator('.grid-host').first().focus()
t = Date.now()
await win.keyboard.press('Control+c')
await win.waitForTimeout(50)
log('Ctrl+C 20000 x 21', (Date.now() - t) + ' ms')
await win.evaluate(() => globalThis.store.setSelection({ ranges: [{ r1: 0, c1: 26, r2: 0, c2: 26 }], active: { r: 0, c: 26 } }))
t = Date.now()
await win.keyboard.press('Control+v')
for (let i = 0; i < 600; i++) { const done = await win.evaluate(() => globalThis.store.getRaw(globalThis.store.active, 19999, 26) !== ''); if (done) break; await win.waitForTimeout(50) }
log('Ctrl+V 20000 x 21', (Date.now() - t) + ' ms')
await win.keyboard.press('Control+z')
await win.waitForTimeout(100)
t = Date.now()
await win.keyboard.press('Control+f')
await win.waitForTimeout(300)
log('open Find dialog', (Date.now() - t) + ' ms')
await win.keyboard.press('Escape')
await frameStats('after operations')
await win.screenshot?.({ path: 'scripts/out/ui-perf.png' }).catch(() => {})
log('heap (renderer)', await win.evaluate(() => (performance.memory ? (performance.memory.usedJSHeapSize / 1048576).toFixed(0) + ' MB' : 'n/a')))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

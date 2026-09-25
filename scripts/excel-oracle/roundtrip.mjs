// Round-trips .xlsx files through the built Cellbook desktop app: copy to out/rt/, open in the app, Ctrl+S.
// Usage: npm run build && node scripts/excel-oracle/roundtrip.mjs file1.xlsx [file2.xlsx ...]
import { _electron as electron } from 'playwright-core'
import fs from 'fs'
import path from 'path'

const outDir = path.resolve('scripts/excel-oracle/out/rt')
fs.mkdirSync(outDir, { recursive: true })
for (const src of process.argv.slice(2)) {
  const dst = path.join(outDir, path.basename(src))
  fs.copyFileSync(src, dst)
  const before = fs.statSync(dst).mtimeMs
  // Fresh profile each run: a leftover AutoRecover backup would otherwise prompt on launch.
  const profile = path.join(outDir, 'profile')
  fs.rmSync(profile, { recursive: true, force: true })
  const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
  // The app pre-loads a hidden spare window for Ctrl+N; use the visible one.
  await app.firstWindow()
  let win
  for (let i = 0; i < 100 && !win; i++) {
    for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p
    if (!win) await new Promise((r) => setTimeout(r, 200))
  }
  win.on('console', (m) => { if (m.type() === 'error') console.log('  console:', m.text()) })
  win.on('pageerror', (e) => console.log('  pageerror:', e.message))
  win.on('dialog', (d) => { console.log('  dialog:', d.message()); void d.accept() })
  await win.waitForFunction(() => globalThis.store)
  await win.evaluate((p) => window.api.openPathRequest(p), dst)
  // The file may land in this window or in the (formerly hidden) spare one: find it by title.
  const stem = path.basename(dst).replace(/\.xlsx$/, '')
  let found
  for (let i = 0; i < 150 && !found; i++) {
    for (const p of app.windows()) if ((await p.title().catch(() => '')).includes(stem)) found = p
    if (!found) await new Promise((r) => setTimeout(r, 200))
  }
  if (found) {
    win = found
    win.on('dialog', (d) => { console.log('  dialog:', d.message()); void d.accept() })
    win.on('pageerror', (e) => console.log('  pageerror:', e.message))
  } else {
    console.log('NOT OPENED', dst, await win.title())
    await win.screenshot({ path: dst + '.png' })
    await app.evaluate(({ app }) => app.exit(0))
    continue
  }
  await win.waitForTimeout(1500)
  const shown = await win.evaluate(() => {
    const out = {}
    for (let r = 0; r < 12; r++) for (let c = 0; c < 16; c++) { const t = store.getDisplay(store.active, r, c).text; if (t) out[String.fromCharCode(65 + c) + (r + 1)] = t }
    return out
  })
  fs.writeFileSync(dst + '.cellbook-display.json', JSON.stringify(shown, null, 1))
  await win.locator('body').click({ position: { x: 300, y: 300 } })
  await win.keyboard.press('Control+s')
  const t0 = Date.now()
  while (fs.statSync(dst).mtimeMs === before && Date.now() - t0 < 30000) await new Promise((r) => setTimeout(r, 200))
  const saved = fs.statSync(dst).mtimeMs !== before
  console.log(saved ? 'saved' : 'NOT SAVED', dst)
  if (!saved) await win.screenshot({ path: dst + '.png' })
  await app.evaluate(({ app }) => app.exit(0))
}

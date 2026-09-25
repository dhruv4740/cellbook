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
  const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.'] })
  const win = await app.firstWindow()
  await win.waitForFunction(() => globalThis.store)
  await app.evaluate(({ BrowserWindow }, [p, bytes]) => {
    const w = BrowserWindow.getAllWindows()[0]
    w.webContents.send('file:openPath', { path: p, data: new Uint8Array(bytes) })
  }, [dst, [...fs.readFileSync(dst)]])
  await win.waitForFunction((name) => document.title.includes(name.replace(/\.xlsx$/, '')), path.basename(dst), { timeout: 30000 })
  await win.waitForTimeout(1500)
  await win.locator('body').click({ position: { x: 300, y: 300 } })
  await win.keyboard.press('Control+s')
  const t0 = Date.now()
  while (fs.statSync(dst).mtimeMs === before && Date.now() - t0 < 30000) await new Promise((r) => setTimeout(r, 200))
  const saved = fs.statSync(dst).mtimeMs !== before
  console.log(saved ? 'saved' : 'NOT SAVED', dst)
  if (!saved) await win.screenshot({ path: dst + '.png' })
  await app.evaluate(({ app }) => app.exit(0))
}

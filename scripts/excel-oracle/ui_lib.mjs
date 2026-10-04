// Shared helpers for the UI drivers: launch the built app, find the data window, read cells, screenshot.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
export async function launch(tag) {
  const profile = path.resolve('scripts/excel-oracle/out/rt/profile-' + tag + '-' + Date.now())
  const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
  await app.firstWindow()
  // keep test windows off-screen and unfocusable so they never steal the user's focus (SHOW=1 to see them)
  if (!process.env.SHOW) await app.evaluate(({ BrowserWindow }) => { for (const w of BrowserWindow.getAllWindows()) { w.setFocusable(false); w.setPosition(-3000, -3000); w.setSize(1920, 1080) } })
  let win
  for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
  await win.waitForFunction(() => globalThis.store)
  const errors = []
  win.on('pageerror', (e) => errors.push(String(e).slice(0, 200)))
  win.on('console', (m) => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
  const api = {
    app, win, errors, profile,
    log: (k, v) => console.log(k.padEnd(34), JSON.stringify(v)),
    set: (r, c, v) => win.evaluate(([r, c, v]) => globalThis.store.setRaw(globalThis.store.active, r, c, v), [r, c, v]),
    raw: (r, c) => win.evaluate(([r, c]) => globalThis.store.getRaw(globalThis.store.active, r, c), [r, c]),
    val: (r, c) => win.evaluate(([r, c]) => globalThis.store.getValue(globalThis.store.active, r, c), [r, c]),
    sel: (r1, c1, r2 = r1, c2 = c1) => win.evaluate(([r1, c1, r2, c2]) => globalThis.store.setSelection({ ranges: [{ r1, c1, r2, c2 }], active: { r: r1, c: c1 } }), [r1, c1, r2, c2]),
    tab: async (name) => { await win.locator('.xl-tab', { hasText: name }).first().click(); await win.waitForTimeout(250) },
    click: async (text) => { await win.getByText(text, { exact: false }).filter({ visible: true }).first().click(); await win.waitForTimeout(250) },
    shot: async (n, region) => {
      const b64 = await app.evaluate(async ({ BrowserWindow }, rect) => {
        let t = BrowserWindow.getAllWindows()[0]
        for (const w of BrowserWindow.getAllWindows()) { if (await w.webContents.executeJavaScript('!!globalThis.store').catch(() => false) && w.isVisible()) t = w }
        return (await (rect ? t.webContents.capturePage(rect) : t.webContents.capturePage())).toPNG().toString('base64')
      }, region)
      fs.mkdirSync('scripts/excel-oracle/out/ui', { recursive: true })
      fs.writeFileSync('scripts/excel-oracle/out/ui/' + n + '.png', Buffer.from(b64, 'base64'))
    },
    close: async () => { await app.close(); fs.rmSync(profile, { recursive: true, force: true }) }
  }
  return api
}

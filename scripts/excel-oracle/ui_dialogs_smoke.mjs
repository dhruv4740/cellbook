// Opens each newer dialog through the app's own state and reports its title plus any renderer errors.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const out = path.resolve('scripts/excel-oracle/out')
const profile = path.resolve(out, 'rt/profile-smoke-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile, '--disable-backgrounding-occluded-windows', '--disable-renderer-backgrounding', '--disable-features=CalculateNativeWinOcclusion'] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store && globalThis.__ui)
const errors = []
win.on('pageerror', (e) => errors.push(String(e)))
win.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()) })
await win.evaluate(() => { const s = globalThis.store; for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) s.setRaw(s.active, r, c, String(r + c)) })
const kinds = [
  ['dataTable'], ['scenarioManager'], ['spelling'], ['sparkline', { type: 'line' }], ['goalSeek'], ['subtotal'], ['pivot'], ['hyperlink']
]
for (const [kind, props] of kinds) {
  await win.evaluate(([k, p]) => globalThis.__ui.setUi({ dialog: { kind: k, props: p } }), [kind, props])
  await win.waitForTimeout(500)
  const title = await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent ?? '(none)')
  console.log(kind.padEnd(16), title)
  await win.evaluate(() => globalThis.__ui.setUi({ dialog: null }))
  await win.waitForTimeout(150)
}
console.log('errors:', errors.length ? errors.join(' || ').slice(0, 800) : 'none')
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-gw-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
for (const tab of ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View']) {
  await win.locator('.xl-tab', { hasText: new RegExp('^' + tab) }).first().click()
  await win.waitForTimeout(400)
  const r = await win.evaluate(() => {
    const host = document.querySelector('.xl-ribbon-groups')
    const groups = [...host.querySelectorAll('.xl-group, .xl-group-collapsed, [class*="xl-group"]')].filter((g) => g.parentElement === host || g.parentElement?.parentElement === host)
    return { avail: Math.round(host.clientWidth), sw: Math.round(host.scrollWidth), groups: [...host.children].map((c) => Math.round(c.getBoundingClientRect().width) + ':' + (c.textContent || '').slice(0, 14)) }
  })
  console.log(tab, JSON.stringify(r))
}
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

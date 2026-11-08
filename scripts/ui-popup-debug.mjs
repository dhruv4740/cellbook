import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-pu-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
win.on('pageerror', (e) => console.log('pageerror', String(e).slice(0, 300)))
win.on('console', (m) => { if (m.type() === 'error') console.log('console.error', m.text().slice(0, 300)) })
await win.evaluate(() => { const s = globalThis.store; s.setRaw(s.active, 0, 0, 'a'); s.setRaw(s.active, 1, 0, 'a') })
await win.locator('.xl-tab', { hasText: 'Data' }).first().click()
await win.waitForTimeout(300)
for (const name of ['Remove Duplicates', 'Consolidate', 'Data Validation', 'Text to Columns']) {
  await win.getByText('Data Tools', { exact: true }).filter({ visible: true }).first().click()
  await win.waitForTimeout(300)
  const info = await win.evaluate((n) => { const b = [...document.querySelectorAll('.xl-group-popup button')].find((x) => x.textContent.includes(n)); return b ? { cls: b.className, disabled: b.disabled, rect: JSON.stringify(b.getBoundingClientRect()) } : null }, name)
  console.log(name, 'button', JSON.stringify(info))
  await win.evaluate((n) => { const b = [...document.querySelectorAll('.xl-group-popup button')].find((x) => x.textContent.includes(n)); b.click() }, name)
  await win.waitForTimeout(300)
  console.log(name, 'after DOM click ->', await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent ?? 'no dialog'))
  await win.keyboard.press('Escape')
  await win.getByText('Data Tools', { exact: true }).filter({ visible: true }).first().click()
  await win.waitForTimeout(300)
  await win.locator('.xl-group-popup').getByText(name).click()
  await win.waitForTimeout(500)
  console.log(name, '->', await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent ?? 'no dialog'))
  await win.keyboard.press('Escape')
  await win.waitForTimeout(200)
}
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

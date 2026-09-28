// Drives the built app with real mouse/keyboard for Flash Fill, Advanced Filter, Goal Seek, Consolidate, manual calculation.
//   node scripts/excel-oracle/ui_features.mjs
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/excel-oracle/out/rt/profile-feat-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) { for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p; if (!win) await new Promise((r) => setTimeout(r, 200)) }
await win.waitForFunction(() => globalThis.store)
const errors = []
win.on('pageerror', (e) => errors.push(String(e).slice(0, 200)))
win.on('console', (m) => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
const log = (k, v) => console.log(k.padEnd(34), JSON.stringify(v))
const get = (r, c) => win.evaluate(([r, c]) => globalThis.store.getRaw(globalThis.store.active, r, c), [r, c])
const val = (r, c) => win.evaluate(([r, c]) => globalThis.store.getValue(globalThis.store.active, r, c), [r, c])
const sel = (r1, c1, r2 = r1, c2 = c1) => win.evaluate(([r1, c1, r2, c2]) => globalThis.store.setSelection({ ranges: [{ r1, c1, r2, c2 }], active: { r: r1, c: c1 } }), [r1, c1, r2, c2])
const tab = async (name) => { await win.locator('.xl-tab', { hasText: name }).first().click(); await win.waitForTimeout(250) }
const click = async (text) => { await win.getByText(text, { exact: false }).filter({ visible: true }).first().click(); await win.waitForTimeout(250) }
// ribbon groups collapse into a popup button at narrow widths: open the group first when the item isn't visible
const clickIn = async (text, group) => {
  const item = win.getByText(text, { exact: false }).filter({ visible: true })
  if (!(await item.count())) { await win.getByText(group, { exact: true }).filter({ visible: true }).first().click(); await win.waitForTimeout(250) }
  await click(text)
}
const shot = async (n) => {
  const b64 = await app.evaluate(async ({ BrowserWindow }) => {
    // the app keeps a spare window: capture the one that holds this test's data
    let target = BrowserWindow.getAllWindows()[0]
    for (const w of BrowserWindow.getAllWindows()) { const has = await w.webContents.executeJavaScript('!!globalThis.store?.getRaw(globalThis.store.active, 0, 0)').catch(() => false); if (has) target = w }
    return (await target.webContents.capturePage()).toPNG().toString('base64')
  })
  fs.mkdirSync('scripts/excel-oracle/out/ui', { recursive: true })
  fs.writeFileSync('scripts/excel-oracle/out/ui/' + n + '.png', Buffer.from(b64, 'base64'))
}
const set = (r, c, v) => win.evaluate(([r, c, v]) => globalThis.store.setRaw(globalThis.store.active, r, c, v), [r, c, v])

// ---- Flash Fill (Ctrl+E) ----
for (const [i, n] of ['John Smith', 'Mary Jane Watson', 'Bob Lee'].entries()) await set(i, 0, n)
await set(0, 1, 'John')
await sel(1, 1)
await win.locator('.grid-host').first().focus()
await win.keyboard.press('Control+e')
await win.waitForTimeout(300)
log('flash fill B2,B3', [await get(1, 1), await get(2, 1)])

// ---- Advanced Filter ----
await win.evaluate(() => { const s = globalThis.store; s.newSheetName; })
const base = 10
const rows = [['Region', 'Sales'], ['East', 10], ['West', 20], ['East', 30], ['North', 5]]
for (const [i, r] of rows.entries()) for (const [j, v] of r.entries()) await set(base + i, j, String(v))
await set(base, 4, 'Region'); await set(base + 1, 4, 'East')
await sel(base, 0)
await tab('Data')
await click('Advanced')
await win.waitForTimeout(300)
const inputs = win.locator('.xl-dialog input[type=text], .xl-dialog input:not([type])')
log('advanced filter dialog inputs', await inputs.count())
await inputs.nth(1).fill('E11:E12')
await win.keyboard.press('Enter')
await win.waitForTimeout(400)
log('hidden rows after adv filter', await win.evaluate(() => Object.keys(globalThis.store.active.layout.hiddenRows)))

// ---- Goal Seek ----
await win.evaluate(() => { const s = globalThis.store; s.clearAllFilters(s.active) })
await set(30, 0, '2'); await set(30, 1, '=A31*3')
await sel(30, 1)
await tab('Data')
await clickIn('What-If Analysis', 'Forecast')
await click('Goal Seek')
await win.waitForTimeout(300)
const gs = win.locator('.xl-dialog input')
log('goal seek inputs', await gs.count())
await gs.nth(1).fill('12')
await gs.nth(2).fill('A31')
await win.keyboard.press('Enter')
await win.waitForTimeout(600)
log('goal seek dialog', await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent))
await win.keyboard.press('Enter')
await win.waitForTimeout(300)
log('goal seek A31 (expect 4)', await val(30, 0))

// ---- Manual calculation ----
await set(40, 0, '1'); await set(40, 1, '=A41*10')
await tab('Formulas')
await clickIn('Calculation Options', 'Calculation')
await click('Manual')
await set(40, 0, '5')
log('manual: B41 stale (expect 10)', await val(40, 1))
log('status bar', await win.evaluate(() => document.querySelector('.xl-statusbar, .xl-status')?.textContent?.slice(0, 40)))
await win.locator('.grid-host').first().focus()
await win.keyboard.press('F9')
await win.waitForTimeout(200)
log('after F9: B41 (expect 50)', await val(40, 1))
await tab('Formulas')
await clickIn('Calculation Options', 'Calculation')
await click('Automatic')

// ---- Consolidate ----
await set(60, 0, 'a'); await set(60, 1, '1'); await set(61, 0, 'b'); await set(61, 1, '2')
await set(60, 3, 'a'); await set(60, 4, '10'); await set(61, 3, 'b'); await set(61, 4, '20')
await sel(70, 0)
await tab('Data')
await win.getByText('Data Tools', { exact: true }).filter({ visible: true }).first().click(); await win.waitForTimeout(300); await shot('datatools-popup'); await win.locator('.xl-group-popup').getByText('Consolidate').click(); await win.waitForTimeout(400)
await win.waitForTimeout(300)
await shot('consolidate-open')
log('dialog title now', await win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent))
const cs = win.locator('.xl-dialog input[type=text], .xl-dialog input:not([type])')
await cs.nth(0).fill('A61:B62')
await win.getByRole('button', { name: 'Add' }).click()
await cs.nth(0).fill('D61:E62')
await win.getByRole('button', { name: 'Add' }).click()
await win.locator('.xl-dialog input[type=checkbox]').nth(1).check()
await win.keyboard.press('Enter')
await win.waitForTimeout(400)
log('consolidate A71:B72', [await get(70, 0), await get(70, 1), await get(71, 0), await get(71, 1)])

log('page errors', errors)
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

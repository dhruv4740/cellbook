// Clicks every ribbon button and every item of every dropdown it opens, resetting a seeded workbook before each click.
// Reports per control: what appeared (dialog title / menu size), renderer errors, or "no visible effect".
//   node scripts/excel-oracle/ui_crawl.mjs [TabName ...]      -> scripts/excel-oracle/out/crawl.json
import fs from 'fs'
import { launch } from './ui_lib.mjs'
const only = process.argv.slice(2)
const t = await launch('crawl')
const { win, app } = t
await app.evaluate(({ dialog, shell }) => {
  shell.openExternal = async () => {}; shell.openPath = async () => ''
  const cancel = async () => ({ canceled: true, filePaths: [], filePath: undefined, response: 1, checkboxChecked: false })
  dialog.showOpenDialog = cancel; dialog.showSaveDialog = cancel; dialog.showMessageBox = cancel; dialog.showMessageBoxSync = () => 1
})
await win.waitForFunction(() => globalThis.__ui)
win.on('dialog', (d) => { d.dismiss().catch(() => {}) })
process.on('unhandledRejection', () => {})
const SKIP = /^(print|share|comments?|sign in|close|minimi|maximi|restore|new window|help|feedback|what.?s new|customize|quick access)/i
const seed = () => win.evaluate(() => {
  const s = globalThis.store
  window.__native = []
  window.alert = (m) => { window.__native.push('alert: ' + m) }; window.confirm = (m) => { window.__native.push('confirm: ' + m); return false }; window.prompt = (m) => { window.__native.push('prompt: ' + m); return null }
  s.newWorkbook()
  globalThis.__ui.setUi({ dialog: null, edit: null, backstage: false, toast: null })
  const sh = s.active
  const rows = [['Name', 'Qty', 'Price', 'Total'], ['apple', 3, 1.5, '=B2*C2'], ['pear', 5, 2.25, '=B3*C3'], ['plum', 2, 4, '=B4*C4'], ['fig', 8, 0.5, '=B5*C5'], ['kiwi', 1, 3, '=B6*C6']]
  rows.forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v))))
  s.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 5, c2: 3 }], active: { r: 1, c: 1 } })
})
const snap = () => win.evaluate(() => {
  const { toast, ...ui } = globalThis.__ui.getUi()
  let h = 0
  const cv = document.querySelector('canvas.xl-canvas')
  if (cv) { try { const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data; for (let i = 0; i < d.length; i += 997) h = (h * 31 + d[i]) | 0 } catch { /* webgl/tainted */ } }
  let u = ''; try { u = JSON.stringify(ui) } catch { u = String(Object.keys(ui)) }
  return JSON.stringify([globalThis.store.version, document.querySelectorAll('.xl-dialog').length, document.querySelectorAll('.xl-menu').length, document.body.innerText.length, u.length, h])
})
const toastText = () => win.evaluate(() => globalThis.__ui.getUi().toast ?? null)
const dialogTitle = () => win.evaluate(() => document.querySelector('.xl-dialog-title')?.textContent ?? (document.querySelector('.xl-dialog') ? '(untitled dialog)' : window.__native?.length ? window.__native[window.__native.length - 1].slice(0, 60) : null))
const closeAll = async () => {
  for (let i = 0; i < 4; i++) {
    const open = await win.evaluate(() => document.querySelectorAll('.xl-dialog, .xl-menu, .xl-color-menu, .xl-group-popup, .xl-combo-list').length)
    if (!open) break
    await win.keyboard.press('Escape'); await win.waitForTimeout(80)
    if (await win.evaluate(() => document.querySelectorAll('.xl-dialog').length)) await win.evaluate(() => globalThis.__ui.setUi({ dialog: null }))
  }
  await win.evaluate(() => globalThis.__ui.setUi({ dialog: null, edit: null, backstage: false }))
}
let curTab = 'Home'
const reset = async () => { await win.evaluate(() => globalThis.__ui?.setUi?.({ toast: null })).catch(() => {}); await seed(); await closeAll(); await win.locator('.xl-tab', { hasText: new RegExp('^' + curTab) }).first().click({ timeout: 3000, force: true }).catch(() => {}); await win.waitForTimeout(120) }
const report = []
const tabs = ['Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Help']
const label = (el) => (el.getAttribute('aria-label') || el.innerText || el.title || '').replace(/\s+/g, ' ').trim()
for (const tabName of tabs) {
  if (only.length && !only.includes(tabName)) continue
  curTab = tabName
  await reset()
  const count = () => win.evaluate(() => document.querySelectorAll('.xl-ribbon-groups button:not([disabled]), .xl-ribbon-groups .xl-combo-caret').length)
  const n = await count()
  console.log(`== ${tabName}: ${n} controls`)
  for (let i = 0; i < n; i++) {
    await reset()
    const btn = win.locator('.xl-ribbon-groups button:not([disabled]), .xl-ribbon-groups .xl-combo-caret').nth(i)
    if (!(await btn.isVisible().catch(() => false))) continue
    const name = (await btn.evaluate((el) => (el.getAttribute('aria-label') || el.innerText || el.title || '').replace(/\s+/g, ' ').trim())).slice(0, 60) || `#${i}`
    const cls = await btn.evaluate((e) => e.className)
    if (SKIP.test(name) || /xl-launcher/.test(cls) && false) { report.push({ tab: tabName, i, name, result: 'skipped' }); continue }
    console.log('>', tabName, i, name)
    const errBefore = t.errors.length
    const before = await snap()
    try { await btn.click({ timeout: 2500 }) } catch (e) { report.push({ tab: tabName, i, name, cls, result: 'click-failed', detail: String(e).slice(0, 900) }); continue }
    await win.waitForTimeout(220)
    const dlg = await dialogTitle()
    const items = await win.evaluate(() => [...document.querySelectorAll('.xl-menu .xl-menu-item, .xl-combo-list .xl-combo-item, .xl-group-popup button, .xl-color-menu .xl-color-swatch, .xl-color-menu .xl-more-colors')].filter((e) => !e.disabled && !e.classList.contains('disabled') && e.offsetParent !== null).map((e) => (e.innerText || e.getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim().slice(0, 50)))
    const entry = { tab: tabName, i, name, cls: cls.slice(0, 40), errors: t.errors.slice(errBefore) }
    const tt = await toastText()
    if (tt && /isn.t available/i.test(tt)) entry.result = 'stub'
    else if (dlg) entry.result = 'dialog: ' + dlg
    else if (items.length) entry.result = `menu: ${items.length} items`
    else entry.result = (await snap()) === before ? 'no visible effect' : 'changed state'
    report.push(entry)
    if (items.length && !dlg && entry.result !== 'stub') {
      entry.items = []
      for (let j = 0; j < items.length; j++) {
        if (SKIP.test(items[j])) { entry.items.push({ j, name: items[j], result: 'skipped' }); continue }
        await reset()
        const e0 = t.errors.length; console.log('  item', j, items[j])
        try {
          await btn.click({ timeout: 2500 }); await win.waitForTimeout(200)
          const cand = win.locator('.xl-menu .xl-menu-item, .xl-combo-list .xl-combo-item, .xl-group-popup button, .xl-color-menu .xl-color-swatch, .xl-color-menu .xl-more-colors').filter({ visible: true })
          const m = await cand.count()
          if (j >= m) { entry.items.push({ j, name: items[j], result: 'vanished' }); continue }
          const b4 = await snap()
          await cand.nth(j).click({ timeout: 2500, force: true }); await win.waitForTimeout(220)
          const d2 = await dialogTitle(); const t2 = await toastText()
          const sub = await win.evaluate(() => document.querySelectorAll('.xl-menu .xl-menu').length)
          entry.items.push({ j, name: items[j], result: t2 && /isn.t available/i.test(t2) ? 'stub' : d2 ? 'dialog: ' + d2 : sub ? 'submenu' : (await snap()) === b4 ? 'no visible effect' : 'changed state', errors: t.errors.slice(e0) })
        } catch (e) { entry.items.push({ j, name: items[j], result: 'click-failed', detail: String(e).slice(0, 900), errors: t.errors.slice(e0) }) }
      }
    }
    if (i % 5 === 4) fs.writeFileSync('scripts/excel-oracle/out/crawl.json', JSON.stringify(report, null, 1))
  }
}
await closeAll()
fs.writeFileSync('scripts/excel-oracle/out/crawl.json', JSON.stringify(report, null, 1))
const bad = report.filter((r) => r.errors?.length || r.items?.some((x) => x.errors?.length))
const stubs = report.filter((r) => r.result === 'stub' || r.items?.some((x) => x.result === 'stub'))
console.log('stubs', report.flatMap((r) => [r.result === 'stub' ? r.tab + ':' + r.name : null, ...(r.items || []).map((x) => x.result === 'stub' ? r.tab + ':' + r.name + '>' + x.name : null)]).filter(Boolean).join(' | '))
const idle = report.filter((r) => r.result === 'no visible effect' || r.items?.some((x) => x.result === 'no visible effect'))
console.log('controls', report.length, '| with errors', bad.length, '| no visible effect', idle.length)
await t.close()

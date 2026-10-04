// Right-clicks a cell, row header, column header and sheet tab; clicks every item of each context menu and reports result.
import fs from 'fs'
import { launch } from './ui_lib.mjs'
const t = await launch('ctx'); const { win, app } = t
await app.evaluate(({ dialog, shell }) => { shell.openExternal = async () => {}; const c = async () => ({ canceled: true, filePaths: [], response: 1 }); dialog.showOpenDialog = c; dialog.showSaveDialog = c; dialog.showMessageBox = c })
win.on('dialog', (d) => d.dismiss().catch(() => {})); process.on('unhandledRejection', () => {})
const targets = { cell: [141, 252], rowHeader: [23, 252], colHeader: [141, 193], sheetTab: [103, 1045] }
const seed = () => win.evaluate(() => { const s = globalThis.store; s.newWorkbook(); globalThis.__ui.setUi({ dialog: null, edit: null, toast: null }); const sh = s.active; [['Name', 'Qty'], ['apple', 3], ['pear', 5], ['plum', 2]].forEach((row, r) => row.forEach((v, c) => s.setRaw(sh, r, c, String(v)))); s.setSelection({ ranges: [{ r1: 1, c1: 1, r2: 1, c2: 1 }], active: { r: 1, c: 1 } }) })
const itemSel = '.xl-menu .xl-menu-item'
const out = []
for (const [name, [x, y]] of Object.entries(targets)) {
  await seed(); await win.mouse.click(700, 600); await win.keyboard.press('Escape')
  await win.mouse.click(x, y, { button: 'right' }); await win.waitForTimeout(250)
  const items = await win.evaluate((s) => [...document.querySelectorAll(s)].filter((e) => e.offsetParent).map((e) => e.innerText.replace(/\s+/g, ' ').trim()), itemSel)
  console.log(name, 'items', items.length); if (!items.length) { out.push({ target: name, result: 'NO MENU' }); continue }
  for (let j = 0; j < items.length; j++) {
    await seed(); await win.keyboard.press('Escape'); await win.mouse.click(x, y, { button: 'right' }); await win.waitForTimeout(200)
    const e0 = t.errors.length
    const it = win.locator(itemSel).filter({ visible: true }).nth(j)
    if (await it.evaluate((e) => e.classList.contains('disabled') || e.disabled).catch(() => true)) { out.push({ target: name, item: items[j], result: 'disabled' }); continue }
    await it.click({ force: true, timeout: 2500 }).catch(() => {}); await win.waitForTimeout(220)
    const r = await win.evaluate(() => { const tt = globalThis.__ui.getUi().toast; return tt && /isn.t available/i.test(tt) ? 'stub' : document.querySelector('.xl-dialog-title')?.textContent ? 'dialog: ' + document.querySelector('.xl-dialog-title').textContent : document.querySelectorAll('.xl-menu .xl-menu').length ? 'submenu' : 'ok' })
    out.push({ target: name, item: items[j], result: r, errors: t.errors.slice(e0) })
    await win.keyboard.press('Escape'); await win.evaluate(() => globalThis.__ui.setUi({ dialog: null }))
  }
}
fs.writeFileSync('scripts/excel-oracle/out/context.json', JSON.stringify(out, null, 1))
for (const o of out) console.log(o.target.padEnd(10), (o.item ?? '').padEnd(34), o.result, o.errors?.length ? 'ERR ' + o.errors : '')
await t.close()

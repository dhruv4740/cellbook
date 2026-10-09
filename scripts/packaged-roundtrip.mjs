// Launches the PACKAGED app (release/win-unpacked), opens a copy of a real workbook, edits one cell, saves with Ctrl+S,
// then compares the saved copy with the original (values, number formats, merges, widths, validation, conditional formats).
//   node scripts/packaged-roundtrip.mjs "samples/private/DESPATCH DHRUV.xlsx"
import { _electron as electron } from 'playwright-core'
import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
const src = path.resolve(process.argv[2])
const dir = path.resolve('scripts/out/rt-packaged'); fs.mkdirSync(dir, { recursive: true })
const copy = path.join(dir, 'copy.xlsx'); fs.copyFileSync(src, copy)
const app = await electron.launch({ executablePath: path.resolve('release/win-unpacked/Cellbook.exe'), args: ['--user-data-dir=' + path.join(dir, 'profile-' + Date.now())] })
let win = await app.firstWindow()
await win.waitForFunction(() => globalThis.store, null, { timeout: 30000 })
const errors = []; win.on('pageerror', (e) => errors.push(String(e)))
await win.evaluate((p) => window.api.openPathRequest(p), copy)
let data
for (let i = 0; i < 300 && !data; i++) {
  for (const p of app.windows()) { const n = await p.evaluate(() => { const u = globalThis.store?.usedRange?.(globalThis.store.active); return u ? u.r2 + 1 : 0 }).catch(() => 0); if (n > 100) data = p }
  if (!data) await new Promise((r) => setTimeout(r, 100))
}
if (!data) { console.log('FAIL file did not open'); await app.close(); process.exit(1) }
win = data
console.log('opened rows', await win.evaluate(() => globalThis.store.usedRange(globalThis.store.active).r2 + 1))
// edit one cell far from the data the comparison reads closely (last column, row 3) then save in place
await win.evaluate(() => { const s = globalThis.store; s.setCells(s.active, [{ r: 2, c: 30, raw: 'EDIT-TEST' }], 'edit') })
await win.keyboard.press('Control+S'); await win.waitForTimeout(3000)
await app.close()
const a = new ExcelJS.Workbook(); await a.xlsx.readFile(src)
const b = new ExcelJS.Workbook(); await b.xlsx.readFile(copy)
let diffs = 0; const note = (m) => { diffs++; if (diffs < 25) console.log('DIFF', m) }
if (a.worksheets.length !== b.worksheets.length) note('sheet count ' + a.worksheets.length + ' vs ' + b.worksheets.length)
a.worksheets.forEach((wa, i) => {
  const wb = b.worksheets[i]; if (!wb) return
  if (wa.name !== wb.name) note('name ' + wa.name + ' vs ' + wb.name)
  wa.eachRow({ includeEmpty: false }, (row, r) => row.eachCell({ includeEmpty: false }, (cell, c) => {
    const o = wb.getCell(r, c)
    const va = JSON.stringify(cell.value), vb = JSON.stringify(o.value)
    if (va !== vb && !(cell.type === o.type && String(cell.value) === String(o.value))) note(`${wa.name}!R${r}C${c} value ${va} vs ${vb}`)
    if ((cell.numFmt || 'General') !== (o.numFmt || 'General')) note(`${wa.name}!R${r}C${c} numFmt ${cell.numFmt} vs ${o.numFmt}`)
  }))
  const ma = Object.keys(wa._merges ?? {}).sort().join(), mb = Object.keys(wb._merges ?? {}).sort().join()
  if (ma !== mb) note(`${wa.name} merges ${ma} vs ${mb}`)
  for (let c = 1; c <= wa.columnCount; c++) { const x = wa.getColumn(c).width, y = wb.getColumn(c).width; if (x && Math.abs((x ?? 0) - (y ?? 0)) > 0.6) note(`${wa.name} col ${c} width ${x} vs ${y}`) }
  const da = Object.keys(wa.dataValidations?.model ?? {}).length, db = Object.keys(wb.dataValidations?.model ?? {}).length
  if (da !== db) note(`${wa.name} data validations ${da} vs ${db}`)
  const ca = (wa.conditionalFormattings ?? []).length, cb = (wb.conditionalFormattings ?? []).length
  if (ca !== cb) note(`${wa.name} conditional formats ${ca} vs ${cb}`)
})
console.log('page errors', errors.length ? errors : 'none')
console.log(diffs ? `${diffs} differences` : 'NO DIFFERENCES (except the deliberate edit cell, which is outside the original used range)')

// Desktop behaviour of the built app: Ctrl+S on a file Excel has open, AutoSave on/off, crash recovery.
// Run from desktop_check.py (which holds the file open in Excel), or alone: node scripts/excel-oracle/desktop_check.mjs <file.xlsx> [skip-locked]
import { _electron as electron } from 'playwright-core'
import fs from 'fs'
import path from 'path'

const file = path.resolve(process.argv[2])
const out = path.resolve('scripts/excel-oracle/out/rt')
const log = (...a) => console.log(...a)

async function launch(profile) {
  const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
  await app.firstWindow()
  const dialogs = []
  let win
  for (let i = 0; i < 100 && !win; i++) {
    for (const p of app.windows()) {
      if (!p.__hooked) { p.__hooked = true; p.on('dialog', (d) => { dialogs.push(d.message()); void d.accept() }) }
      if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p
    }
    if (!win) await new Promise((r) => setTimeout(r, 200))
  }
  await win.waitForFunction(() => globalThis.store)
  return { app, win, dialogs }
}
async function openFile(s, p) {
  await s.win.evaluate((p) => window.api.openPathRequest(p), p)
  const stem = path.basename(p).replace(/\.xlsx$/, '')
  for (let i = 0; i < 150; i++) {
    for (const w of s.app.windows()) {
      if ((await w.title().catch(() => '')).includes(stem)) {
        if (!w.__hooked) { w.__hooked = true; w.on('dialog', (d) => { s.dialogs.push(d.message()); void d.accept() }) }
        s.win = w; return
      }
    }
    await new Promise((r) => setTimeout(r, 200))
  }
  throw new Error('file did not open')
}
async function typeAt(win, ref, text) {
  await win.evaluate(() => store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } }))
  await win.locator('.grid-host').first().focus()
  await win.keyboard.press('Control+g').catch(() => {})
  await win.keyboard.press('Escape')
  await win.locator('.grid-host').first().focus()
  await win.keyboard.type(text)
  await win.keyboard.press('Enter')
}

// 1. Ctrl+S while Excel has the file open
if (process.argv[3] !== 'skip-locked') {
  const s = await launch(path.join(out, 'profile-desk-' + Date.now()))
  await openFile(s, file)
  const before = fs.readFileSync(file)
  await typeAt(s.win, 'A1', 'edited-while-open')
  await s.win.keyboard.press('Control+s')
  await s.win.waitForTimeout(3000)
  const after = fs.readFileSync(file)
  log('1. locked save: dialogs =', JSON.stringify(s.dialogs))
  log('   file on disk unchanged:', before.equals(after), '| edit still in app:', await s.win.evaluate(() => store.getRaw(store.active, 0, 0)), '| still marked unsaved:', await s.win.evaluate(() => store.dirty))
  log('   temp files left next to it:', fs.readdirSync(path.dirname(file)).filter((f) => f.includes(path.basename(file)) && f !== path.basename(file)))
  await s.app.evaluate(({ app }) => app.exit(0))
  process.stdout.write('LOCKED-DONE\n')
  await new Promise((r) => process.stdin.once('data', r)) // python releases Excel, then continues us
}

// 2. AutoSave on/off
if (process.argv[4] !== 'crash-only') {
  const s = await launch(path.join(out, 'profile-desk-' + Date.now()))
  await openFile(s, file)
  await s.win.locator('.xl-autosave').click()
  await s.win.waitForTimeout(500)
  const m0 = fs.statSync(file).mtimeMs
  await typeAt(s.win, 'A1', 'autosaved')
  await s.win.waitForTimeout(6000)
  const m1 = fs.statSync(file).mtimeMs
  log('2. AutoSave on: saved within 6 s:', m1 !== m0, '| dirty after:', await s.win.evaluate(() => store.dirty), '| label:', await s.win.locator('.xl-autosave').innerText())
  await s.win.locator('.xl-autosave').click()
  await s.win.waitForTimeout(500)
  await typeAt(s.win, 'A1', 'not-autosaved')
  await s.win.waitForTimeout(6000)
  log('   AutoSave off: file untouched for 6 s:', fs.statSync(file).mtimeMs === m1, '| label:', await s.win.locator('.xl-autosave').innerText(), '| dialogs:', JSON.stringify(s.dialogs))
  await s.win.evaluate(() => { store.dirty = false })
  await s.app.evaluate(({ app }) => app.exit(0))
}

// 3. Crash recovery: unsaved edits, kill the process, relaunch with the same profile
{
  const profile = path.join(out, 'profile-crash-' + Date.now())
  const s = await launch(profile)
  await typeAt(s.win, 'A1', 'unsaved-before-crash')
  log('3. waiting 65 s for the AutoRecover backup...')
  await s.win.waitForTimeout(65000)
  const { execSync } = await import('child_process')
  execSync('taskkill /PID ' + s.app.process().pid + ' /T /F') // like Task Manager > End task
  await new Promise((r) => setTimeout(r, 5000))
  const s2 = await launch(profile)
  await s2.win.waitForTimeout(4000)
  const wins = []
  for (const w of s2.app.windows()) wins.push([await w.title().catch(() => ''), await w.evaluate(() => [store.getRaw(store.active, 0, 0), store.dirty]).catch(() => null)])
  log('   after relaunch: dialogs =', JSON.stringify(s2.dialogs), '| windows =', JSON.stringify(wins))
  await s2.app.evaluate(({ app }) => app.exit(0))
}
process.exit(0)

// Line-based driver for the built Cellbook app, used by clipboard_check.py. Commands on stdin:
//   select A1 | select A1:C3 | key Control+v | set A1 <raw> | dump A1:D8 | quit
// Each command answers one JSON line on stdout.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import readline from 'readline'

const profile = path.resolve('scripts/excel-oracle/out/rt/profile-clip-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
let win
for (let i = 0; i < 100 && !win; i++) {
  for (const p of app.windows()) if (await p.evaluate(() => document.visibilityState === 'visible').catch(() => false)) win = p
  if (!win) await new Promise((r) => setTimeout(r, 200))
}
await win.waitForFunction(() => globalThis.store)
win.on('dialog', (d) => { console.error('dialog:', d.message()); void d.accept() })
const cell = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: +m[2] - 1, c: c - 1 } }
const out = (o) => process.stdout.write(JSON.stringify(o) + '\n')
out({ ready: true })
for await (const line of readline.createInterface({ input: process.stdin })) {
  const [cmd, arg, ...rest] = line.split(' ')
  try {
    if (cmd === 'select') {
      const [a, b = a] = arg.split(':').map(cell)
      await win.evaluate(([a, b]) => store.setSelection({ ranges: [{ r1: a.r, c1: a.c, r2: b.r, c2: b.c }], active: { r: a.r, c: a.c } }), [a, b])
      await win.bringToFront(); await win.locator('.grid-host').first().focus()
      out({ ok: true })
    } else if (cmd === 'key') {
      await win.keyboard.press(arg)
      await win.waitForTimeout(800)
      out({ ok: true })
    } else if (cmd === 'set') {
      const a = cell(arg)
      await win.evaluate(([a, raw]) => store.setRaw(store.active, a.r, a.c, raw), [a, rest.join(' ').replace(/\\n/g, '\n')])
      out({ ok: true })
    } else if (cmd === 'dump') {
      const [a, b] = arg.split(':').map(cell)
      out(await win.evaluate(([a, b]) => {
        const rows = []
        for (let r = a.r; r <= b.r; r++) {
          const row = []
          for (let c = a.c; c <= b.c; c++) { const v = store.getValue(store.active, r, c); row.push({ raw: store.getRaw(store.active, r, c), value: v, type: typeof v, text: store.getDisplay(store.active, r, c).text, fmt: store.getStyle(store.active, r, c)?.numFmt }) }
          rows.push(row)
        }
        return rows
      }, [a, b]))
    } else if (cmd === 'quit') {
      out({ ok: true })
      break
    }
  } catch (e) { out({ error: String(e) }) }
}
await app.evaluate(({ app }) => app.exit(0)).catch(() => {})
process.exit(0)

// Times Electron's clipboard write/read in the main process for big payloads.
import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-clip-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
for (const rows of [1000, 8000, 20000]) {
  const res = await app.evaluate(async ({ clipboard, ClipboardItem }, rows) => {
    const cell = 'value-123456'
    const line = Array.from({ length: 21 }, () => cell).join('\t')
    const text = Array.from({ length: rows }, () => line).join('\n')
    const html = '<table>' + Array.from({ length: rows }, () => '<tr>' + Array.from({ length: 21 }, () => `<td style='font-weight:bold'>${cell}</td>`).join('') + '</tr>').join('') + '</table>'
    const out = { mb: (text.length / 1048576).toFixed(1) + ' MB text, ' + (html.length / 1048576).toFixed(1) + ' MB html' }
    let t = performance.now()
    clipboard.write({ text })
    out.writeTextOnly = Math.round(performance.now() - t)
    t = performance.now()
    clipboard.write({ text, html })
    out.writeTextHtml = Math.round(performance.now() - t)
    t = performance.now()
    await clipboard.write([new ClipboardItem({ 'text/plain': text, 'text/html': html })])
    out.writeClipboardItem = Math.round(performance.now() - t)
    t = performance.now()
    const back = clipboard.readText()
    out.readText = Math.round(performance.now() - t) + ' (' + back.length + ' chars)'
    return out
  }, rows)
  console.log(rows, JSON.stringify(res))
}
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

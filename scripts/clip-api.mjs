import { _electron as electron } from 'playwright-core'
import path from 'path'
import fs from 'fs'
const profile = path.resolve('scripts/out/profile-clipapi-' + Date.now())
const app = await electron.launch({ executablePath: 'node_modules/electron/dist/electron.exe', args: ['.', '--user-data-dir=' + profile] })
await app.firstWindow()
const res = await app.evaluate(async ({ clipboard, ClipboardItem }) => {
  const out = {}
  const desc = (v) => (v && typeof v.then === 'function' ? 'Promise' : typeof v) + ':' + String(v).slice(0, 30)
  try { const r = clipboard.write({ text: 'legacy-text', html: '<b>x</b>' }); out.legacyWrite = desc(r) } catch (e) { out.legacyWriteErr = String(e).slice(0, 120) }
  try { out.readAfterLegacy = desc(await clipboard.readText()) } catch (e) { out.readAfterLegacyErr = String(e).slice(0, 120) }
  try { const r = clipboard.write([new ClipboardItem({ 'text/plain': 'item-text', 'text/html': '<i>y</i>' })]); out.itemWrite = desc(r); await r } catch (e) { out.itemWriteErr = String(e).slice(0, 160) }
  try { out.readAfterItem = desc(await clipboard.readText()) } catch (e) { out.readAfterItemErr = String(e).slice(0, 120) }
  try { const r = clipboard.write([new ClipboardItem({ 'text/plain': new Blob(['blob-text'], { type: 'text/plain' }) })]); await r; out.readAfterBlob = desc(await clipboard.readText()) } catch (e) { out.blobErr = String(e).slice(0, 160) }
  try { out.readHtml = desc(await clipboard.readHTML?.()) } catch (e) { out.readHtmlErr = String(e).slice(0, 120) }
  return out
})
console.log(JSON.stringify(res, null, 1))
await app.close()
fs.rmSync(profile, { recursive: true, force: true })

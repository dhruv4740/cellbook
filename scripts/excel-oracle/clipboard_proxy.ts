// Helper for clipboard_proxy.py: pastes Excel's tab-separated text into Cellbook, and writes the HTML
// Cellbook puts on the clipboard for a typical block.
import fs from 'fs'
let written: { text: string; html?: string } = { text: '' }
const excelText = fs.readFileSync(process.argv[2], 'utf16le').replace(/^﻿/, '')
;(globalThis as any).window = { api: { clipboardReadText: async () => excelText, clipboardWrite: async (d: any) => { written = d } } }
const { store } = await import('../../src/renderer/src/model/store.ts')
const cmds = await import('../../src/renderer/src/grid/commands.ts')
const s = store.active
store.setSelection({ ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 } })
await cmds.paste()
console.log('--- Excel text pasted into Cellbook:', JSON.stringify(excelText))
for (let c = 0; c < 10; c++) console.log('  ', String.fromCharCode(65 + c) + '1', JSON.stringify(store.getRaw(s, 0, c)), JSON.stringify(store.getValue(s, 0, c)), JSON.stringify(store.getDisplay(s, 0, c).text), store.getStyle(s, 0, c)?.numFmt ?? '')
for (const [c, raw] of ['1234567.891', '3/9/2026', 'a, b, c', 'Line1\nLine2', "'00012", '-12.5%', 'TRUE', '=A5*2', 'x"y', '₹5,000'].entries()) store.setRaw(s, 4, c, raw)
store.setStyle?.(s, { r1: 4, c1: 0, r2: 4, c2: 0 }, { numFmt: '#,##0.00' })
store.setSelection({ ranges: [{ r1: 4, c1: 0, r2: 4, c2: 9 }], active: { r: 4, c: 0 } })
await cmds.copy()
console.log('--- Cellbook copy text:', JSON.stringify(written.text))
fs.writeFileSync(process.argv[3], '<html><head><meta charset="utf-8"></head><body>' + (written.html ?? '') + '</body></html>')

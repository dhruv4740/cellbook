// Loads the formula workbook into Cellbook's engine and writes each formula's value as JSON.
// Run: npx vite-node scripts/excel-oracle/cellbook_values.ts
import fs from 'fs'
const { parseXlsx } = await import('../../src/renderer/src/io/fastXlsx.ts')
const { store } = await import('../../src/renderer/src/model/store.ts')
store.load(parseXlsx(new Uint8Array(fs.readFileSync(process.argv[2] ?? 'scripts/excel-oracle/out/formulas.xlsx'))) as any, 'fx', null)
const out: any[] = []
for (let r = 9; r < 9 + 170; r++) { const raw = store.getRaw(store.active, r, 0); if (!raw) continue; const v = store.getValue(store.active, r, 0); out.push([r + 1, raw, v && typeof v === 'object' ? (v as any).error : v, store.getDisplay(store.active, r, 0).text]) }
fs.writeFileSync(process.argv[3] ?? 'scripts/excel-oracle/out/cellbook-values.json', JSON.stringify(out))

// Thin CLI over the app's xlsx reader/writer, for the Excel round-trip checks (feature_rt.py).
//   npx vite-node scripts/excel-oracle/feature_io.ts read  <in.xlsx>            -> JSON on stdout (layout features per sheet)
//   npx vite-node scripts/excel-oracle/feature_io.ts write <in.json> <out.xlsx> -> WorkbookData JSON in, xlsx out
import fs from 'fs'
import { parseXlsx } from '../../src/renderer/src/io/fastXlsx'
import { writeXlsx } from '../../src/renderer/src/io/xlsx'

const [mode, a, b] = process.argv.slice(2)
if (mode === 'read') {
  const d = parseXlsx(new Uint8Array(fs.readFileSync(a)))
  const out = { protection: d.protection, calcMode: d.calcMode, names: d.names, sheets: d.sheets.map((s) => ({ name: s.name, layout: s.layout, cells: s.cells })) }
  process.stdout.write(JSON.stringify(out))
} else if (mode === 'write') {
  const data = JSON.parse(fs.readFileSync(a, 'utf-8'))
  fs.writeFileSync(b, await writeXlsx(data))
} else {
  console.error('usage: read <xlsx> | write <json> <xlsx>')
  process.exit(2)
}

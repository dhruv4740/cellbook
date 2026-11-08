// Prints how many pages Cellbook's paginator makes for each sheet of a workbook JSON (compared with Excel in feature_rt.py pages).
import fs from 'fs'
import { store } from '../../src/renderer/src/model/store'
import { layoutSheet } from '../../src/renderer/src/io/print'

store.load(JSON.parse(fs.readFileSync(process.argv[2], 'utf-8')))
const out: Record<string, unknown> = {}
for (const sh of store.sheets) {
  const sp = layoutSheet(sh, 'active')
  out[sh.name] = sp ? { pages: sp.pages.length, scale: sp.scale, hbreaks: [...new Set(sp.pages.map((p) => p.r1))].filter((r) => r > sp.rg.r1) } : null
}
process.stdout.write(JSON.stringify(out))

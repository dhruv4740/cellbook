// Cellbook's values for a sample of rows of samples/private/scaled-<n>.xlsx (formulas computed by the engine).
import fs from 'fs'
import { parseXlsx } from '../../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../../src/renderer/src/model/store'
const file = process.argv[2]
const s = new WorkbookStore()
const t0 = performance.now()
s.load(parseXlsx(new Uint8Array(fs.readFileSync(file))), 'x', null)
const load = performance.now() - t0
const sh = s.active
const n = s.usedRange(sh)!.r2 + 1
const rows: number[] = []
let seed = 12345
for (let i = 0; i < 300; i++) { seed = (seed * 1103515245 + 12345) & 0x7fffffff; rows.push(1 + (seed % (Math.min(n, 20000) - 1))) }
const out: Record<string, unknown> = { n, loadMs: load, rows: {} }
for (const r of rows) (out.rows as Record<number, unknown>)[r] = { w: s.getValue(sh, r, 22), d: s.getValue(sh, r, 3), e: s.getValue(sh, r, 4) }
process.stdout.write(JSON.stringify(out))

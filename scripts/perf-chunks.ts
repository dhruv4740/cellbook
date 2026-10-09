// Experiment: how should a 50k x 21 block be written into HyperFormula? One call, chunks, suspended evaluation, no batch.
import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
import { WorkbookStore } from '../src/renderer/src/model/store'
const n = Number(process.argv[2] ?? 50000)
const src = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
const base = src.sheets[0]
const cells: typeof base.cells = {}
for (let r = 0; r < n; r++) { const sr = r === 0 ? 0 : 1 + ((r - 1) % 520); for (let c = 0; c < 21; c++) { const x = base.cells[sr + ',' + c]; if (x) cells[r + ',' + c] = x } }
const s = new WorkbookStore()
s.load({ ...src, sheets: [{ ...base, cells }] }, 'x', undefined)
const id = s.hfId(s.active)
const m = s.hf.getSheetSerialized(id) as unknown[][]
const rev = [m[0], ...m.slice(1).reverse()]
const time = (label: string, fn: () => void): void => { const t = performance.now(); fn(); console.log(label.padEnd(40), (performance.now() - t).toFixed(0), 'ms') }
const put = (rows: unknown[][], chunk: number): void => { for (let i = 0; i < rows.length; i += chunk) s.hf.setCellContents({ sheet: id, row: i, col: 0 }, rows.slice(i, i + chunk) as never) }
time('one call', () => put(rev, rev.length))
time('one call back', () => put(m, m.length))
time('chunks 1000 (no batch)', () => put(rev, 1000))
time('chunks 1000 back (no batch)', () => put(m, 1000))
time('chunks 1000 in batch', () => s.hf.batch(() => put(rev, 1000)))
time('chunks 1000 in batch back', () => s.hf.batch(() => put(m, 1000)))
time('suspend + one call + resume', () => { s.hf.suspendEvaluation(); put(rev, rev.length); s.hf.resumeEvaluation() })
time('suspend + one call back + resume', () => { s.hf.suspendEvaluation(); put(m, m.length); s.hf.resumeEvaluation() })
time('suspend + chunks 1000 + resume', () => { s.hf.suspendEvaluation(); put(rev, 1000); s.hf.resumeEvaluation() })

// Cellbook side of func_oracle.py: loads func_cases.json and prints every formula's value as JSON.
import fs from 'fs'
import { WorkbookStore } from '../../src/renderer/src/model/store'
const cases = JSON.parse(fs.readFileSync('scripts/excel-oracle/func_cases.json', 'utf-8'))
const s = new WorkbookStore()
const sh = s.active
cases.data.forEach((row: unknown[], r: number) => row.forEach((v, c) => {
  if (v === null) return
  s.setRaw(sh, r, c, typeof v === 'string' ? (v === '' ? "'" : /^[-+\d.=<>]|^true$|^false$/i.test(v) ? "'" + v : v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v))
}))
const out = cases.formulas.map((f: string, i: number) => {
  s.setRaw(sh, i, 10, '=' + f)
  const v = s.getValue(sh, i, 10)
  return typeof v === 'object' && v !== null ? (v as { error: string }).error : v
})
process.stdout.write(JSON.stringify(out))

import fs from 'fs'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
const d = parseXlsx(new Uint8Array(fs.readFileSync('samples/private/DESPATCH DHRUV.xlsx')))
for (const s of d.sheets) {
  const keys = Object.keys(s.cells)
  let maxR = 0, maxC = 0, f = 0
  for (const k of keys) { const [r, c] = k.split(',').map(Number); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); if (s.cells[k].raw.startsWith('=')) f++ }
  console.log(s.name, 'cells', keys.length, 'rows', maxR + 1, 'cols', maxC + 1, 'formulas', f, 'merges', s.layout?.merges?.length, 'af', !!s.layout?.autoFilter, 'cf', s.layout?.conditionalFormats?.length)
  for (let r = 0; r < 4; r++) console.log('  ', Array.from({ length: Math.min(maxC + 1, 14) }, (_, c) => (s.cells[r + ',' + c]?.raw ?? '').toString().slice(0, 14)).join(' | '))
}

import fs from 'fs'
import v8 from 'v8'
import { parseXlsx } from '../src/renderer/src/io/fastXlsx'
const bytes = new Uint8Array(fs.readFileSync(process.argv[2] ?? 'samples/private/scaled-50000.xlsx'))
let t = performance.now()
const data = parseXlsx(bytes)
console.log('parseXlsx', (performance.now() - t).toFixed(0), 'ms')
t = performance.now()
const c = structuredClone(data)
console.log('structuredClone of result', (performance.now() - t).toFixed(0), 'ms', Object.keys(c.sheets[0].cells).length, 'cells')
t = performance.now()
const s = v8.serialize(data)
console.log('v8.serialize', (performance.now() - t).toFixed(0), 'ms', (s.length / 1048576).toFixed(0), 'MB')

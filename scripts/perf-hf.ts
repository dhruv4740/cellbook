import { HyperFormula } from 'hyperformula'
const n = Number(process.argv[2] ?? 5000)
const mk = (opts: object) => {
  const rows: (string | number)[][] = []
  for (let r = 0; r < n; r++) rows.push(['k' + (r % 700), r, `=VLOOKUP(A${r + 1},$A$1:$B$${n},2,FALSE)`])
  const t = performance.now()
  const hf = HyperFormula.buildFromArray(rows, { licenseKey: 'gpl-v3', ...opts })
  return { hf, ms: performance.now() - t }
}
for (const [name, o] of [['default', {}], ['colIndex', { useColumnIndex: true }], ['colIndex+stats', { useColumnIndex: true, useStats: true }]] as const) {
  const { ms } = mk(o)
  console.log(name.padEnd(16), 'build', ms.toFixed(0))
}
// n distinct-key VLOOKUP into sorted data

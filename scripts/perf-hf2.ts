import { HyperFormula } from 'hyperformula'
import { registerExtraFunctions } from '../src/renderer/src/engine/extraFunctions'
import { patchHyperFormulaGraph, instrumentHf } from '../src/renderer/src/engine/hfPatches'
registerExtraFunctions(); patchHyperFormulaGraph()
const n = Number(process.argv[2] ?? 5000)
const which = process.argv[3] ?? 'xlookup'
const f: Record<string, (r: number) => string> = {
  vlookup: (r) => `=VLOOKUP(A${r + 1},$A$1:$B$${n},2,FALSE())`,
  xlookup: (r) => `=XLOOKUP(A${r + 1},$A$1:$A$${n},$B$1:$B$${n},"na")`,
  match: (r) => `=MATCH(A${r + 1},$A$1:$A$${n},0)`,
  sumifs: (r) => `=SUMIFS($B$1:$B$${n},$A$1:$A$${n},A${r + 1})`,
  countif: (r) => `=COUNTIF($A$1:$A$${n},A${r + 1})`,
  sumif: (r) => `=SUMIF($A$1:$A$${n},A${r + 1},$B$1:$B$${n})`,
  index: (r) => `=INDEX($B$1:$B$${n},MATCH(A${r + 1},$A$1:$A$${n},0))`
}
const rows: (string | number)[][] = []
for (let r = 0; r < n; r++) rows.push(['k' + (r % 700), r, f[which](r)])
let t = performance.now()
const hf = instrumentHf(HyperFormula.buildFromArray(rows, { licenseKey: 'gpl-v3', useColumnIndex: true }))
console.log(which.padEnd(8), 'build', (performance.now() - t).toFixed(0))
t = performance.now(); hf.setCellContents({ sheet: 0, row: 3, col: 1 }, 9999); console.log('         edit', (performance.now() - t).toFixed(0))

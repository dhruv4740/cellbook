import { HyperFormula } from 'hyperformula'
const n = Number(process.argv[2] ?? 5000)
const patch = process.argv[3] === 'patch'
if (patch) {
  const probe: any = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
  const proto = Object.getPrototypeOf(probe.dependencyGraph)
  const orig = proto.processCellDependencies
  proto.processCellDependencies = function (cellDependencies: any[], endVertex: any) {
    const rest: any[] = []
    for (const dep of cellDependencies) {
      if (dep && dep.start && dep.end && typeof dep.isFinite === 'function') {
        const rv = this.getRange(dep.start, dep.end)
        if (rv !== undefined && rv.idInGraph !== undefined && this.graph.hasNode(rv)) { this.graph.addEdge(rv.idInGraph, endVertex.idInGraph); continue }
      }
      rest.push(dep)
    }
    if (rest.length) orig.call(this, rest, endVertex)
  }
}
const rows: (string | number)[][] = []
for (let r = 0; r < n; r++) rows.push(['k' + (r % 700), r, `=VLOOKUP(A${r + 1},$A$1:$B$${n},2,FALSE())`])
let t = performance.now()
const hf = HyperFormula.buildFromArray(rows, { licenseKey: 'gpl-v3', useColumnIndex: true })
console.log(patch ? 'patched' : 'stock', 'build', (performance.now() - t).toFixed(0), 'sample', hf.getCellValue({ sheet: 0, row: 5, col: 2 }))
t = performance.now(); hf.setCellContents({ sheet: 0, row: 3, col: 0 }, 'zz'); console.log(' edit key', (performance.now() - t).toFixed(0))
t = performance.now(); hf.setCellContents({ sheet: 0, row: 3, col: 1 }, 9999); console.log(' edit val', (performance.now() - t).toFixed(0), hf.getCellValue({ sheet: 0, row: 3 + 700, col: 2 }))

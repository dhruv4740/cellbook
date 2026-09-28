import { WorkbookStore } from '../src/renderer/src/model/store'
const s = new WorkbookStore()
const sh = s.active
for (const [i, raw] of ["'007", '007', "'abc", 'abc', '1/2/2024', "'1/2/2024", '=1+1', '5.50', "'5"].entries()) {
  s.setRaw(sh, i, 0, raw)
  const ser = s.hf.getCellSerialized({ sheet: 0, row: i, col: 0 })
  console.log(JSON.stringify(raw).padEnd(14), 'serialized', JSON.stringify(ser), typeof ser, '| getRaw', JSON.stringify(s.getRaw(sh, i, 0)), '| value', JSON.stringify(s.getValue(sh, i, 0)))
}
console.log('matrix', JSON.stringify(s.hf.getSheetSerialized(0)))

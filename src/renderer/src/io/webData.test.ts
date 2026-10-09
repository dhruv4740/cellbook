import { describe, expect, it } from 'vitest'
import { parseWebTables } from './webData'

describe('parseWebTables', () => {
  it('reads HTML tables with captions, entities and colspan', () => {
    const t = parseWebTables('<p>x</p><table><caption>Sales</caption><tr><th>A</th><th colspan=2>B &amp; C</th></tr><tr><td>1</td><td>2<td>3</table><table><tr><td>z</td></tr></table>', 'text/html')
    expect(t.map((x) => x.name)).toEqual(['Sales', 'Table 2'])
    expect(t[0].rows).toEqual([['A', 'B & C', ''], ['1', '2', '3']])
  })
  it('reads a JSON array of records', () => {
    expect(parseWebTables('[{"a":1,"b":"x"},{"a":2,"c":true}]', 'application/json')[0].rows).toEqual([['a', 'b', 'c'], ['1', 'x', ''], ['2', '', 'true']])
  })
  it('reads CSV', () => {
    expect(parseWebTables('a,"b,c"\n1,2\n', 'text/csv')[0].rows).toEqual([['a', 'b,c'], ['1', '2']])
  })
})

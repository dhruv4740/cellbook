import { describe, it, expect } from 'vitest'
import { moveRefs } from './moveRefs'
const src = { r1: 0, c1: 0, r2: 1, c2: 0 } // A1:A2
describe('moveRefs', () => {
  it('follows moved cells', () => {
    expect(moveRefs('=A1+B1', 'S', 'S', src, 8, 2, 'S')).toBe('=C9+B1')
    expect(moveRefs('=$A$2*2', 'S', 'S', src, 1, 0, 'S')).toBe('=$A$3*2')
    expect(moveRefs('=SUM(A1:A2)', 'S', 'S', src, 0, 1, 'S')).toBe('=SUM(B1:B2)')
    expect(moveRefs('=SUM(A1:A3)', 'S', 'S', src, 0, 1, 'S')).toBe('=SUM(A1:A3)')
    expect(moveRefs('="A1"&A1', 'S', 'S', src, 1, 1, 'S')).toBe('="A1"&B2')
    expect(moveRefs('=S!A1', 'T', 'S', src, 0, 1, 'S')).toBe('=S!B1')
    expect(moveRefs('=A1', 'T', 'S', src, 0, 1, 'S')).toBe('=A1')
    expect(moveRefs('=A1', 'S', 'S', src, 0, 0, 'Other Sheet')).toBe("='Other Sheet'!A1")
    expect(moveRefs('=LOG10(A1)', 'S', 'S', src, 1, 0, 'S')).toBe('=LOG10(A2)')
  })
})

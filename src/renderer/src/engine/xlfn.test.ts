import { describe, it, expect } from 'vitest'
import { addFuturePrefixes, stripFuturePrefixes } from './xlfn'

describe('_xlfn prefixes', () => {
  it('adds them for newer functions only, never inside strings', () => {
    expect(addFuturePrefixes('XLOOKUP(A1,B:B,C:C)+SUM(1)')).toBe('_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(1)')
    expect(addFuturePrefixes('IFS(A1>1,"IFS(",TRUE,MAXIFS(B:B,C:C,">0"))')).toBe('_xlfn.IFS(A1>1,"IFS(",TRUE,_xlfn.MAXIFS(B:B,C:C,">0"))')
    expect(addFuturePrefixes('FILTER(A1:A9,B1:B9>0)')).toBe('_xlfn._xlws.FILTER(A1:A9,B1:B9>0)')
    expect(addFuturePrefixes('STDEV.S(A:A)*VLOOKUP(1,A:B,2,0)')).toBe('_xlfn.STDEV.S(A:A)*VLOOKUP(1,A:B,2,0)')
  })
  it('strips them on load and round-trips', () => {
    const f = '_xlfn.XLOOKUP(A1,B:B,C:C,_xlfn.CONCAT("a","_xlfn."))'
    expect(stripFuturePrefixes(f)).toBe('XLOOKUP(A1,B:B,C:C,CONCAT("a","_xlfn."))')
    expect(addFuturePrefixes(stripFuturePrefixes(f))).toBe(f)
    expect(stripFuturePrefixes('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)')).toBe('LET(x,1,x+1)')
  })
})

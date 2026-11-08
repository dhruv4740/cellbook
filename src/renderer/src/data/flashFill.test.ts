import { describe, expect, it } from 'vitest'
import { flashFill } from './flashFill'

const col = (xs: string[]): string[][] => xs.map((x) => [x])

describe('flash fill', () => {
  it('extracts a first name', () => {
    const r = flashFill(col(['John Smith', 'Mary Jane Watson', 'Bob Lee']), ['John', '', ''])
    expect(r).toEqual([null, 'Mary', 'Bob'])
  })
  it('extracts the last word', () => {
    const r = flashFill(col(['John Smith', 'Mary Jane Watson', 'Bob Lee']), ['Smith', '', ''])
    expect(r).toEqual([null, 'Watson', 'Lee'])
  })
  it('rearranges with a constant', () => {
    const r = flashFill(col(['John Smith', 'Mary Watson', 'Bob Lee']), ['Smith, John', '', ''])
    expect(r).toEqual([null, 'Watson, Mary', 'Lee, Bob'])
  })
  it('changes case', () => {
    expect(flashFill(col(['abc', 'def']), ['ABC', ''])).toEqual([null, 'DEF'])
  })
  it('builds from two columns', () => {
    const r = flashFill([['John', 'Smith'], ['Mary', 'Lee']], ['J. Smith', ''])
    expect(r).toEqual([null, 'M. Lee'])
  })
  it('extracts digits from mixed text', () => {
    const r = flashFill(col(['ID-1234-X', 'ID-99-Y', 'ID-5-Z']), ['1234', '', ''])
    expect(r).toEqual([null, '99', '5'])
  })
  it('skips a header example', () => {
    const r = flashFill(col(['Full Name', 'John Smith', 'Mary Lee']), ['First', 'John', ''])
    expect(r).toEqual([null, null, 'Mary'])
  })
  it('gives up when there is no pattern', () => {
    expect(flashFill(col(['a', 'b', 'c', 'd']), ['zzz', 'qqq', 'xxx', ''])).toBeNull()
  })
})

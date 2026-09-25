import { describe, it, expect } from 'vitest'
import { adjustDecimals } from './decimals'

describe('Increase / Decrease Decimal', () => {
  it('General starts from what is shown', () => {
    expect(adjustDecimals('General', '3.14159', 1)).toBe('0.000000')
    expect(adjustDecimals('General', '3.14159', -1)).toBe('0.0000')
    expect(adjustDecimals(undefined, '12', 1)).toBe('0.0')
    expect(adjustDecimals(undefined, '12', -1)).toBe('0')
  })
  it('keeps currency, grouping, colours and sections', () => {
    expect(adjustDecimals('₹#,##0.00', '', 1)).toBe('₹#,##0.000')
    expect(adjustDecimals('₹#,##0.00', '', -1)).toBe('₹#,##0.0')
    expect(adjustDecimals('#,##0.0', '', -1)).toBe('#,##0')
    expect(adjustDecimals('#,##0', '', 1)).toBe('#,##0.0')
    expect(adjustDecimals('0.00;[Red]-0.00', '', 1)).toBe('0.000;[Red]-0.000')
    expect(adjustDecimals('0%', '', 1)).toBe('0.0%')
    expect(adjustDecimals('"Rs "0.00', '', 1)).toBe('"Rs "0.000')
    expect(adjustDecimals('0.00E+00', '', -1)).toBe('0.0E+00')
    expect(adjustDecimals('0E+00', '', 1)).toBe('0.0E+00')
  })
  it('leaves dates and text alone', () => {
    expect(adjustDecimals('dd-mm-yyyy', '', 1)).toBe('dd-mm-yyyy')
    expect(adjustDecimals('@', '', 1)).toBe('@')
  })
})

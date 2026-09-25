import { describe, it, expect, beforeEach } from 'vitest'
import { formatValue, detectInputFormat, buildNumberFormat, adjustDecimals, FORMAT_CATEGORIES } from './numfmt'
import { setLakhGrouping, setDayFirst } from '../engine/locale'

beforeEach(() => { setLakhGrouping(false); setDayFirst(false) })

describe('digit grouping follows the PC region (Excel on English (India))', () => {
  it('groups lakhs and crores', () => {
    setLakhGrouping(true)
    expect(formatValue(1234567.891, '#,##0.00').text).toBe('12,34,567.89')
    expect(formatValue(123456789, '#,##0').text).toBe('12,34,56,789')
    expect(formatValue(-1234, '#,##0').text).toBe('-1,234')
    expect(formatValue(999, '#,##0').text).toBe('999')
  })
})

function t(value: number | string | boolean | null, fmt?: string) {
  return formatValue(value as never, fmt).text
}

describe('General format', () => {
  it('integers', () => {
    expect(t(0)).toBe('0')
    expect(t(1234)).toBe('1234')
    expect(t(-1234)).toBe('-1234')
  })
  it('decimals trim trailing zeros', () => {
    expect(t(1.5)).toBe('1.5')
    expect(t(0.1 + 0.2)).toBe('0.3')
    expect(t(1.10)).toBe('1.1')
  })
  it('up to 11 significant chars, switches to scientific for very large', () => {
    expect(t(123456789012)).toBe('1.23457E+11')
  })
  it('booleans centered TRUE/FALSE', () => {
    expect(formatValue(true as never)).toEqual({ text: 'TRUE', align: 'center' })
    expect(formatValue(false as never)).toEqual({ text: 'FALSE', align: 'center' })
  })
  it('errors centered as-is', () => {
    expect(formatValue({ error: '#DIV/0!' } as never)).toEqual({ text: '#DIV/0!', align: 'center' })
  })
  it('strings left aligned, unmodified', () => {
    expect(formatValue('hello' as never)).toEqual({ text: 'hello', align: 'left' })
  })
  it('numbers align right', () => {
    expect(formatValue(5 as never).align).toBe('right')
  })
})

describe('Number formats', () => {
  it('0', () => { expect(t(1234.5678, '0')).toBe('1235') })
  it('0.00', () => { expect(t(1234.5, '0.00')).toBe('1234.50') })
  it('#,##0', () => { expect(t(1234567, '#,##0')).toBe('1,234,567') })
  it('#,##0.00', () => { expect(t(1234.5, '#,##0.00')).toBe('1,234.50') })
  it('rounds half away from zero', () => {
    expect(t(1.005, '0.00')).toBe('1.01')
    expect(t(-1.005, '0.00')).toBe('-1.01')
    expect(t(2.5, '0')).toBe('3')
    expect(t(-2.5, '0')).toBe('-3')
  })
  it('negative single-section gets minus', () => {
    expect(t(-5, '0.00')).toBe('-5.00')
  })
  it('two-section: negative uses 2nd section, no auto sign', () => {
    expect(t(-1234, '#,##0;(#,##0)')).toBe('(1,234)')
    expect(t(1234, '#,##0;(#,##0)')).toBe('1,234')
  })
  it('three-section: zero uses 3rd', () => {
    expect(t(0, '0.00;(0.00);"zero"')).toBe('zero')
    expect(t(5, '0.00;(0.00);"zero"')).toBe('5.00')
    expect(t(-5, '0.00;(0.00);"zero"')).toBe('(5.00)')
  })
  it('color tags', () => {
    expect(formatValue(-5 as never, '0;[Red]-0').color).toBe('#FF0000')
    expect(formatValue(5 as never, '[Blue]0').color).toBe('#0000FF')
  })
  it('conditions', () => {
    expect(t(150, '[>100]"High";[<=100]"Low"')).toBe('High')
    expect(t(50, '[>100]"High";[<=100]"Low"')).toBe('Low')
  })
  it('# suppresses leading/trailing zeros', () => {
    expect(t(0.5, '#.##')).toBe('.5')
    expect(t(5, '#')).toBe('5')
  })
  it('0 pattern with no value shows blank-ish per Excel (all # zero)', () => {
    expect(t(0, '#')).toBe('')
  })
  it('trailing comma scaling', () => {
    expect(t(12000000, '#,##0,,')).toBe('12')
    expect(t(12345, '#,##0,')).toBe('12')
  })
  it('percent', () => {
    expect(t(0.156, '0%')).toBe('16%')
    expect(t(0.1525, '0.00%')).toBe('15.25%')
  })
  it('quoted literal and escape', () => {
    expect(t(5, '0" units"')).toBe('5 units')
    expect(t(5, '0\\%')).toBe('5%')
  })
})

describe('Currency & Accounting', () => {
  it('$#,##0.00', () => { expect(t(1234.5, '$#,##0.00')).toBe('$1,234.50') })
  it('$#,##0_);($#,##0)', () => {
    expect(t(1234, '$#,##0_);($#,##0)')).toBe('$1,234 ')
    expect(t(-1234, '$#,##0_);($#,##0)')).toBe('($1,234)')
  })
  it('currency red negative', () => {
    const r = formatValue(-1234 as never, '$#,##0.00_);[Red]($#,##0.00)')
    expect(r.text).toBe('($1,234.00)')
    expect(r.color).toBe('#FF0000')
  })
  it('accounting format', () => {
    const code = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    expect(t(1234.5, code)).toBe(' $ 1,234.50 ')
    expect(t(-1234.5, code)).toBe(' $ (1,234.50)')
    expect(t(0, code)).toBe(' $ -   ')
  })
})

describe('Scientific', () => {
  it('0.00E+00', () => { expect(t(12345, '0.00E+00')).toBe('1.23E+04') })
  it('##0.0E+0', () => { expect(t(12345, '##0.0E+0')).toBe('12.3E+3') })
})

describe('Fraction', () => {
  it('# ?/?', () => { expect(t(1.5, '# ?/?')).toBe('1 1/2') })
  it('# ??/??', () => { expect(t(0.02, '# ??/??')).toBe(' 1/50') })
  it('fixed denom # ?/8', () => { expect(t(1.4, '# ?/8')).toBe('1 3/8') })
  it('whole number shows no fraction part', () => { expect(t(3, '# ?/?')).toBe('3') })
})

describe('Dates', () => {
  it('serial 1 = 1900-01-01', () => { expect(t(1, 'm/d/yyyy')).toBe('1/1/1900') })
  it('serial 60 = fake 1900-02-29', () => { expect(t(60, 'm/d/yyyy')).toBe('2/29/1900') })
  it('serial 61 = 1900-03-01', () => { expect(t(61, 'm/d/yyyy')).toBe('3/1/1900') })
  it('serial 0 = 1900-01-00', () => { expect(t(0, 'm/d/yyyy')).toBe('1/0/1900') })
  it('m/d/yyyy known date', () => {
    // 2024-01-15 -> serial 45306
    expect(t(45306, 'm/d/yyyy')).toBe('1/15/2024')
  })
  it('d-mmm-yy', () => { expect(t(45306, 'd-mmm-yy')).toBe('15-Jan-24') })
  it('d-mmm', () => { expect(t(45306, 'd-mmm')).toBe('15-Jan') })
  it('mmm-yy', () => { expect(t(45306, 'mmm-yy')).toBe('Jan-24') })
  it('dddd, mmmm d, yyyy', () => {
    // 2024-01-15 is a Monday
    expect(t(45306, 'dddd, mmmm d, yyyy')).toBe('Monday, January 15, 2024')
  })
  it('mmmmm single-letter month', () => { expect(t(45306, 'mmmmm')).toBe('J') })
})

describe('Times', () => {
  it('h:mm AM/PM', () => {
    // 0.5 = 12:00 PM
    expect(t(0.5, 'h:mm AM/PM')).toBe('12:00 PM')
    expect(t(0, 'h:mm AM/PM')).toBe('12:00 AM')
  })
  it('h:mm:ss AM/PM', () => {
    const serial = (13 * 3600 + 5 * 60 + 9) / 86400
    expect(t(serial, 'h:mm:ss AM/PM')).toBe('1:05:09 PM')
  })
  it('h:mm 24h', () => {
    const serial = (13 * 3600 + 5 * 60) / 86400
    expect(t(serial, 'h:mm')).toBe('13:05')
  })
  it('mm:ss', () => {
    const serial = (5 * 60 + 9) / 86400
    expect(t(serial, 'mm:ss')).toBe('05:09')
  })
  it('[h]:mm:ss elapsed over 24h', () => {
    const serial = 1 + (13 * 3600 + 5 * 60 + 9) / 86400 // 1 day + 13:05:09
    expect(t(serial, '[h]:mm:ss')).toBe('37:05:09')
  })
  it('mm:ss.0 fractional seconds', () => {
    const serial = (5 * 60 + 9.4) / 86400
    expect(t(serial, 'mm:ss.0')).toBe('05:09.4')
  })
  it('m/d/yyyy h:mm combined', () => {
    const serial = 45306 + 0.5
    expect(t(serial, 'm/d/yyyy h:mm')).toBe('1/15/2024 12:00')
  })
})

describe('Long date/time built-ins', () => {
  it('long date', () => {
    expect(t(45306, '[$-x-sysdate]dddd, mmmm d, yyyy')).toBe('Monday, January 15, 2024')
  })
  it('long time', () => {
    const serial = (13 * 3600 + 5 * 60 + 9) / 86400
    expect(t(serial, '[$-x-systime]h:mm:ss AM/PM')).toBe('1:05:09 PM')
  })
})

describe('Text format', () => {
  it('@ passthrough', () => { expect(t('abc', '@')).toBe('abc') })
  it('@ with literal', () => { expect(t('abc', '"Item: "@')).toBe('Item: abc') })
  it('4-section text uses 4th section', () => {
    expect(t('x', '0;0;0;"[" @ "]"')).toBe('[ x ]')
  })
})

describe('detectInputFormat', () => {
  it('$1,234.50', () => { expect(detectInputFormat('$1,234.50')).toBe('"$"#,##0.00') })
  it('$5', () => { expect(detectInputFormat('$5')).toBe('$#,##0') })
  it('15%', () => { expect(detectInputFormat('15%')).toBe('0%') })
  it('12.5%', () => { expect(detectInputFormat('12.5%')).toBe('0.00%') })
  it('1,234', () => { expect(detectInputFormat('1,234')).toBe('#,##0') })
  it('1,234.5', () => { expect(detectInputFormat('1,234.5')).toBe('#,##0.00') })
  it('1/31/2025', () => { expect(detectInputFormat('1/31/2025')).toBe('m/d/yyyy') })
  it('1/31', () => { expect(detectInputFormat('1/31')).toBe('d-mmm') })
  it('Jan 2025', () => { expect(detectInputFormat('Jan 2025')).toBe('mmm-yy') })
  it('31-Jan', () => { expect(detectInputFormat('31-Jan')).toBe('d-mmm') })
  it('10:30', () => { expect(detectInputFormat('10:30')).toBe('h:mm') })
  it('10:30 PM', () => { expect(detectInputFormat('10:30 PM')).toBe('h:mm AM/PM') })
  it('10:30:15', () => { expect(detectInputFormat('10:30:15')).toBe('h:mm:ss') })
  it('1.5E+3', () => { expect(detectInputFormat('1.5E+3')).toBe('0.00E+00') })
  it('plain numbers -> undefined', () => { expect(detectInputFormat('1234')).toBeUndefined() })
  it('formulas -> undefined', () => { expect(detectInputFormat('=SUM(A1:A2)')).toBeUndefined() })
  it('text -> undefined', () => { expect(detectInputFormat('hello world')).toBeUndefined() })
})

describe('FORMAT_CATEGORIES', () => {
  it('has all categories', () => {
    for (const cat of ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Special', 'Custom']) {
      expect(FORMAT_CATEGORIES[cat]).toBeDefined()
      expect(FORMAT_CATEGORIES[cat].length).toBeGreaterThan(0)
    }
  })
  it('Date category codes are all valid, renderable format strings', () => {
    for (const entry of FORMAT_CATEGORIES.Date) {
      expect(() => formatValue(45306 as never, entry.code)).not.toThrow()
    }
  })
})

describe('buildNumberFormat / adjustDecimals', () => {
  it('builds a basic number format', () => {
    expect(buildNumberFormat({ decimals: 2, thousands: true, negative: 0 })).toBe('#,##0.00;-#,##0.00')
  })
  it('builds currency with parens+red', () => {
    expect(buildNumberFormat({ decimals: 2, thousands: true, negative: 3, symbol: '$' })).toBe('"$"#,##0.00;[Red]("$"#,##0.00)')
  })
  it('increases decimals on plain format', () => {
    expect(adjustDecimals('0.00', 1)).toBe('0.000')
  })
  it('decreases decimals on plain format', () => {
    expect(adjustDecimals('0.00', -1)).toBe('0.0')
  })
  it('decreases to no decimal point', () => {
    expect(adjustDecimals('0.0', -1)).toBe('0')
  })
  it('increases from no decimal point', () => {
    expect(adjustDecimals('#,##0', 1)).toBe('#,##0.0')
  })
  it('adjusts General using value context', () => {
    expect(adjustDecimals('General', 1, 1.5)).toBe('0.00')
  })
})

describe('Special formats', () => {
  it('zip code', () => { expect(t(90210, '00000')).toBe('90210') })
  it('zip+4', () => { expect(t(902101234, '00000-0000')).toBe('90210-1234') })
  it('SSN', () => { expect(t(123456789, '000-00-0000')).toBe('123-45-6789') })
})

describe('Caching', () => {
  it('parses a format code only once and reuses it', () => {
    const code = '#,##0.00'
    const a = formatValue(1 as never, code)
    const b = formatValue(2 as never, code)
    expect(a.text).toBe('1.00')
    expect(b.text).toBe('2.00')
  })
  it('accepts an optional colWidthPx opts arg without affecting output', () => {
    expect(formatValue(1234 as never, '#,##0', { colWidthPx: 80 }).text).toBe('1,234')
  })
})

describe('More number rounding / edge cases', () => {
  it('handles negative zero cleanly', () => { expect(t(-0.001, '0.00')).toBe('0.00') })
  it('large integer general', () => { expect(t(99999999999)).toBe('99999999999') })
  it('very small general uses scientific', () => { expect(t(0.0000000001)).toContain('E-') })
  it('zero with 0.00 shows 0.00', () => { expect(t(0, '0.00')).toBe('0.00') })
  it('#,##0.00 with zero', () => { expect(t(0, '#,##0.00')).toBe('0.00') })
})

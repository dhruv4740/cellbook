// Date entry follows the PC's regional order, like Excel: on a day-first locale (India, UK, ...)
// "3/9/2026" is 3 September, on a US locale it is March 9.
function detectDayFirst(): boolean {
  try {
    const parts = new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'numeric', year: 'numeric' }).formatToParts(new Date(2020, 11, 31))
    const order = parts.filter((p) => p.type === 'day' || p.type === 'month').map((p) => p.type)
    return order[0] === 'day'
  } catch {
    return false
  }
}

export const DAY_FIRST = detectDayFirst()

export const DATE_FORMATS = DAY_FIRST
  ? ['DD/MM/YYYY', 'DD/MM/YY', 'YYYY-MM-DD', 'YYYY/MM/DD']
  : ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY-MM-DD', 'YYYY/MM/DD']

/** Number format given to a date typed as d/m/y (or m/d/y), matching Excel's regional short date. */
export const SHORT_DATE_FORMAT = DAY_FIRST ? 'dd-mm-yyyy' : 'm/d/yyyy'

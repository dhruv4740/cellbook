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

// Excel groups "#,##0" with the PC's digit grouping: on English (India) 1234567 shows as 12,34,567.
function detectLakhGrouping(): boolean {
  try {
    return new Intl.NumberFormat(undefined, { useGrouping: true }).format(1234567).replace(/\D/g, ',') === '12,34,567'
  } catch {
    return false
  }
}

export let LAKH_GROUPING = detectLakhGrouping()

export function setLakhGrouping(on: boolean): void {
  LAKH_GROUPING = on
}

const dateFormatsFor = (dayFirst: boolean): string[] => dayFirst
  ? ['DD/MM/YYYY', 'DD/MM/YY', 'YYYY-MM-DD', 'YYYY/MM/DD']
  : ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY-MM-DD', 'YYYY/MM/DD']

export const DATE_FORMATS = dateFormatsFor(DAY_FIRST)

/** Number format given to a date typed as d/m/y (or m/d/y), matching Excel's regional short date. */
export let SHORT_DATE_FORMAT = DAY_FIRST ? 'dd-mm-yyyy' : 'm/d/yyyy'

/** For tests: pretend the PC uses a US (false) or day-first (true) date order. Affects stores created afterwards. */
export function setDayFirst(dayFirst: boolean): void {
  DATE_FORMATS.splice(0, DATE_FORMATS.length, ...dateFormatsFor(dayFirst))
  SHORT_DATE_FORMAT = dayFirst ? 'dd-mm-yyyy' : 'm/d/yyyy'
}

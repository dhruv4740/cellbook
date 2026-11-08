// Date entry follows the PC's regional order, like Excel: on a day-first locale (India, UK, ...)
// "3/9/2026" is 3 September, on a US locale it is March 9.
// In the desktop app the page's own locale is the Windows display language, so use the region the
// main process passes in; in tests/the web build fall back to the environment's default.
const REGION: string | undefined = (globalThis as { api?: { regionLocale?: string } }).api?.regionLocale || undefined

function detectDayFirst(): boolean {
  try {
    const parts = new Intl.DateTimeFormat(REGION, { day: 'numeric', month: 'numeric', year: 'numeric' }).formatToParts(new Date(2020, 11, 31))
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
    return new Intl.NumberFormat(REGION, { useGrouping: true }).format(1234567).replace(/\D/g, ',') === '12,34,567'
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

// The PC's currency symbol for Home > Currency / Accounting (₹ on English (India), $ on US, £ on UK...).
function detectCurrency(): string {
  try {
    const region = (REGION ?? Intl.DateTimeFormat().resolvedOptions().locale).split('-').pop()!.toUpperCase()
    const map: Record<string, string> = { IN: '₹', US: '$', GB: '£', AE: 'AED', AU: '$', CA: '$', SG: '$', NZ: '$', EU: '€', DE: '€', FR: '€', IE: '€', NL: '€', ES: '€', IT: '€', JP: '¥', CN: '¥', PK: 'Rs', LK: 'Rs', NP: 'Rs', BD: '৳', ZA: 'R' }
    return map[region] ?? '$'
  } catch {
    return '$'
  }
}
export const CURRENCY_SYMBOL = detectCurrency()
const q = (sym: string) => (sym.length === 1 ? sym : `"${sym} "`)
/** Home > Number Format "Currency" for this PC, e.g. ₹#,##0.00 */
export const CURRENCY_FORMAT = `${q(CURRENCY_SYMBOL)}#,##0.00`
/** Home > Number Format "Accounting" for this PC. */
export const ACCOUNTING_FORMAT = `_ ${q(CURRENCY_SYMBOL)}* #,##0.00_ ;_ ${q(CURRENCY_SYMBOL)}* -#,##0.00_ ;_ ${q(CURRENCY_SYMBOL)}* "-"??_ ;_ @_ `
/** Long date as Excel shows it for the region. */
export const LONG_DATE_FORMAT = DAY_FIRST ? 'dd mmmm yyyy' : 'dddd, mmmm d, yyyy'

// Home > Fill > Series...: the value sequence Excel's Series dialog produces.
export type SeriesType = 'linear' | 'growth' | 'date'
export type DateUnit = 'day' | 'weekday' | 'month' | 'year'

/** Excel date serial -> JS UTC date and back (1900 system, serial 60 = the fake 29-Feb-1900). */
function toDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000)
}
function toSerial(d: Date): number {
  return Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 86400000)
}

/**
 * Values after `start` for `count` further cells (or fewer, when `stop` is passed first).
 * Linear adds `step`, growth multiplies by it, date steps by `step` units of `unit`.
 */
export function seriesValues(start: number, count: number, type: SeriesType, step: number, stop?: number, unit: DateUnit = 'day'): number[] {
  const out: number[] = []
  let cur = start
  const past = (v: number) => stop !== undefined && (step >= 0 ? v > stop + 1e-12 : v < stop - 1e-12)
  for (let k = 1; k <= count; k++) {
    if (type === 'linear') cur = start + k * step
    else if (type === 'growth') cur = cur * step
    else {
      const d = toDate(start)
      if (unit === 'day') cur = start + k * step
      else if (unit === 'weekday') {
        let v = toDate(cur)
        let left = Math.abs(step)
        const dir = step < 0 ? -1 : 1
        while (left > 0) { v = new Date(v.getTime() + dir * 86400000); const wd = v.getUTCDay(); if (wd !== 0 && wd !== 6) left-- }
        cur = toSerial(v)
      } else {
        const months = unit === 'month' ? k * step : k * step * 12
        const y = d.getUTCFullYear(), m = d.getUTCMonth() + months, day = d.getUTCDate()
        const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate()
        cur = toSerial(new Date(Date.UTC(y, m, Math.min(day, last))))
      }
    }
    if (past(cur)) break
    out.push(type === 'linear' ? +cur.toPrecision(15) : cur)
  }
  return out
}

// Excel's fill-handle series logic: given the raw text of the seed cells (in drag order) and how
// many more cells to produce, returns the raw text for each new cell. Pure/testable; commands.ts
// wires it to the store (formulas are shifted separately, by the caller, with shiftFormula).
const WEEKDAYS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
const WEEKDAYS_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const MONTHS_LONG = [
  'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'
]

function matchCase(sample: string, word: string): string {
  if (sample === sample.toUpperCase()) return word.toUpperCase()
  if (sample[0] === sample[0]?.toLowerCase() && sample !== sample.toLowerCase()) return word
  if (sample === sample.toLowerCase()) return word.toLowerCase()
  return word
}

function listIndex(list: string[], s: string): number {
  return list.findIndex((w) => w.toLowerCase() === s.trim().toLowerCase())
}

// mm/dd/yyyy or m/d/yyyy, matching Excel's default US date entry.
const DATE_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/

function parseDate(s: string): Date | null {
  const m = DATE_RE.exec(s.trim())
  if (!m) return null
  let [, mo, da, yr] = m
  let year = +yr
  if (yr.length === 2) year += year >= 30 ? 1900 : 2000
  const d = new Date(year, +mo - 1, +da)
  if (d.getMonth() !== +mo - 1) return null // rolled over -> invalid (e.g. 2/30)
  return d
}
function fmtDate(d: Date): string {
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

function trailingNumber(s: string): { prefix: string; digits: string } | null {
  const m = /^(.*?)(\d+)$/.exec(s)
  if (!m) return null
  return { prefix: m[1], digits: m[2] }
}

function parseNum(s: string): number | null {
  if (s.trim() === '' || isNaN(Number(s))) return null
  return Number(s)
}

function linearStep(nums: number[]): number {
  if (nums.length === 1) return 1
  // OLS slope over evenly spaced integer x = 0..n-1.
  const n = nums.length
  const xBar = (n - 1) / 2
  const yBar = nums.reduce((a, b) => a + b, 0) / n
  let num = 0, den = 0
  nums.forEach((y, x) => {
    num += (x - xBar) * (y - yBar)
    den += (x - xBar) * (x - xBar)
  })
  return den === 0 ? 0 : num / den
}

/** Generates `count` more raw values continuing the seed series (drag order, seed[0] closest to selection). */
export function generateSeries(seed: string[], count: number, copyOnly = false): string[] {
  const out: string[] = []
  if (seed.length === 0) return out

  if (!copyOnly) {
    // 1. Weekday / month name cycling.
    for (const list of [WEEKDAYS_SHORT, WEEKDAYS_LONG, MONTHS_SHORT, MONTHS_LONG]) {
      const idxs = seed.map((s) => listIndex(list, s))
      if (idxs.every((i) => i >= 0)) {
        const step = idxs.length >= 2 ? idxs[1] - idxs[0] : 1
        for (let i = 0; i < count; i++) {
          const idx = (((idxs[idxs.length - 1] + step * (i + 1)) % list.length) + list.length) % list.length
          out.push(matchCase(seed[0], list[idx]))
        }
        return out
      }
    }

    // 2. Dates: increment by day (or the seeds' own day-step).
    const dates = seed.map(parseDate)
    if (dates.every((d) => d !== null)) {
      const days = (dates as Date[]).map((d) => d.getTime())
      const stepMs = days.length >= 2 ? days[1] - days[0] : 86400000
      const last = days[days.length - 1]
      for (let i = 0; i < count; i++) out.push(fmtDate(new Date(last + stepMs * (i + 1))))
      return out
    }

    // 3. Text with a trailing number ('Item 1' -> 'Item 2').
    const tn = seed.map(trailingNumber)
    if (tn.every((t) => t !== null) && tn.some((t) => t!.prefix !== '')) {
      const prefixesMatch = tn.every((t) => t!.prefix === tn[0]!.prefix)
      if (prefixesMatch) {
        const nums = tn.map((t) => +t!.digits)
        const step = nums.length >= 2 ? nums[1] - nums[0] : 1
        const width = tn[tn.length - 1]!.digits.length
        const last = nums[nums.length - 1]
        for (let i = 0; i < count; i++) {
          const v = last + step * (i + 1)
          const digits = String(Math.max(0, v))
          out.push(tn[0]!.prefix + (tn[0]!.digits[0] === '0' ? digits.padStart(width, '0') : digits))
        }
        return out
      }
    }

    // 4. Pure numbers: linear trend (2+ seeds); a single number just copies.
    const nums = seed.map(parseNum)
    if (nums.every((n) => n !== null) && seed.length >= 2) {
      const vals = nums as number[]
      const step = linearStep(vals)
      const last = vals[vals.length - 1]
      for (let i = 0; i < count; i++) out.push(String(round11(last + step * (i + 1))))
      return out
    }
  }

  // 5. Fallback: cycle the seed values as-is (plain copy).
  for (let i = 0; i < count; i++) out.push(seed[i % seed.length])
  return out
}

function round11(n: number): number {
  return +n.toPrecision(11)
}

import { WorkbookStore, SheetState } from '../model/store'
import { CellStyle, CfRule, Range, inRange, normRange, isError } from '../model/types'
import { shiftFormula } from '../engine/shiftFormula'

export interface CfEvaluation {
  style?: Partial<CellStyle>
  dataBar?: { fraction: number; color: string; negative?: boolean }
  colorScale?: string
  icon?: { set: string; index: number }
}

interface CacheEntry {
  version: number
  evaluations: Map<string, CfEvaluation> // key: 'r,c'
}

const sheetCache = new Map<number, CacheEntry>()

/** Invalidates cache for a sheet or all sheets. */
export function invalidateCfCache(sheetUid?: number): void {
  if (sheetUid !== undefined) sheetCache.delete(sheetUid)
  else sheetCache.clear()
}

/** Parses hex color to RGB. */
function hexToRgb(hex: string): { r: number; g: number; b: number } {
  let clean = hex.replace('#', '')
  if (clean.length === 3) clean = clean.split('').map((c) => c + c).join('')
  const n = parseInt(clean, 16)
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 }
}

/** Interpolates two hex colors by factor t (0..1). */
export function interpolateColor(c1: string, c2: string, t: number): string {
  const rgb1 = hexToRgb(c1)
  const rgb2 = hexToRgb(c2)
  const r = Math.round(rgb1.r + (rgb2.r - rgb1.r) * t)
  const g = Math.round(rgb1.g + (rgb2.g - rgb1.g) * t)
  const b = Math.round(rgb1.b + (rgb2.b - rgb1.b) * t)
  return '#' + [r, g, b].map((x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0')).join('')
}

/** Computes the full evaluation map for all rules on a sheet. */
function buildSheetEvaluations(store: WorkbookStore, sheet: SheetState): Map<string, CfEvaluation> {
  const result = new Map<string, CfEvaluation>()
  const rules = (sheet.layout.conditionalFormats || []).slice().sort((a, b) => a.priority - b.priority)
  if (!rules.length) return result

  const used = store.usedRange(sheet)
  if (!used) return result

  const todaySerial = Math.floor((Date.now() - new Date(1899, 11, 30).getTime()) / 86400000)

  // Pre-calculate aggregate metrics for rules that need them (top10, aboveAverage, duplicateValues, min/max)
  for (const rule of rules) {
    const cellsInRange: { r: number; c: number; raw: string; val: any }[] = []
    const numbersInRange: number[] = []
    const valueCounts = new Map<string, number>()

    for (const rg of rule.ranges) {
      const n = normRange(rg)
      for (let r = n.r1; r <= n.r2; r++) {
        for (let c = n.c1; c <= n.c2; c++) {
          const raw = store.getRaw(sheet, r, c)
          const val = store.getValue(sheet, r, c)
          if (raw !== '' || val !== null) {
            cellsInRange.push({ r, c, raw, val })
            if (typeof val === 'number') numbersInRange.push(val)
            const k = String(val ?? raw).toLowerCase()
            valueCounts.set(k, (valueCounts.get(k) || 0) + 1)
          }
        }
      }
    }

    let avg = 0
    if (numbersInRange.length > 0) {
      avg = numbersInRange.reduce((a, b) => a + b, 0) / numbersInRange.length
    }
    const minVal = numbersInRange.length ? Math.min(...numbersInRange) : 0
    const maxVal = numbersInRange.length ? Math.max(...numbersInRange) : 0

    // Top / Bottom cutoff
    let topCutoff = -Infinity, bottomCutoff = Infinity
    if (rule.type === 'top10' && numbersInRange.length > 0) {
      const sorted = [...numbersInRange].sort((a, b) => b - a)
      const count = rule.percent
        ? Math.max(1, Math.round((sorted.length * (rule.rank || 10)) / 100))
        : Math.min(sorted.length, Math.max(1, rule.rank || 10))
      if (rule.bottom) {
        bottomCutoff = sorted[sorted.length - count]
      } else {
        topCutoff = sorted[count - 1]
      }
    }

    const baseR = rule.ranges[0]?.r1 ?? 0
    const baseC = rule.ranges[0]?.c1 ?? 0

    // Evaluate cells in rule ranges
    for (const rg of rule.ranges) {
      const n = normRange(rg)
      for (let r = n.r1; r <= n.r2; r++) {
        for (let c = n.c1; c <= n.c2; c++) {
          const key = `${r},${c}`
          const cur = result.get(key)
          // If a previous rule had stopIfTrue and matched, skip
          if (cur && (cur as any)._stopped) continue

          const raw = store.getRaw(sheet, r, c)
          const val = store.getValue(sheet, r, c)

          let matches = false
          let dataBar: CfEvaluation['dataBar']
          let colorScale: string | undefined
          let icon: CfEvaluation['icon']

          switch (rule.type) {
            case 'cellIs': {
              const num = typeof val === 'number' ? val : parseFloat(String(val))
              const op = rule.operator || 'greaterThan'
              const v1 = parseFloat(rule.value || '0')
              const v2 = parseFloat(rule.value2 || '0')

              if (!isNaN(num)) {
                if (op === 'greaterThan') matches = num > v1
                else if (op === 'lessThan') matches = num < v1
                else if (op === 'equal') matches = num === v1
                else if (op === 'notEqual') matches = num !== v1
                else if (op === 'greaterThanOrEqual') matches = num >= v1
                else if (op === 'lessThanOrEqual') matches = num <= v1
                else if (op === 'between') matches = num >= Math.min(v1, v2) && num <= Math.max(v1, v2)
                else if (op === 'notBetween') matches = num < Math.min(v1, v2) || num > Math.max(v1, v2)
              } else if (typeof val === 'string') {
                if (op === 'equal') matches = val.toLowerCase() === (rule.value || '').toLowerCase()
                else if (op === 'notEqual') matches = val.toLowerCase() !== (rule.value || '').toLowerCase()
              }
              break
            }

            case 'textContains': {
              const str = (typeof val === 'string' ? val : store.getDisplay(sheet, r, c).text).toLowerCase()
              matches = str.includes((rule.text || rule.value || '').toLowerCase())
              break
            }

            case 'dateOccurring': {
              if (typeof val === 'number') {
                const diff = val - todaySerial
                const d = rule.date || 'today'
                if (d === 'yesterday') matches = diff === -1
                else if (d === 'today') matches = diff === 0
                else if (d === 'tomorrow') matches = diff === 1
                else if (d === 'inTheLast7Days') matches = diff >= -7 && diff <= 0
                else if (d === 'thisWeek') matches = diff >= -3 && diff <= 3
                else if (d === 'lastWeek') matches = diff >= -10 && diff < -3
                else if (d === 'nextWeek') matches = diff > 3 && diff <= 10
                else if (d === 'thisMonth') matches = diff >= -15 && diff <= 15
                else if (d === 'lastMonth') matches = diff < -15 && diff >= -45
                else if (d === 'nextMonth') matches = diff > 15 && diff <= 45
              }
              break
            }

            case 'duplicateValues': {
              const k = String(val ?? raw).toLowerCase()
              matches = (valueCounts.get(k) || 0) > 1
              break
            }

            case 'uniqueValues': {
              const k = String(val ?? raw).toLowerCase()
              matches = (valueCounts.get(k) || 0) === 1
              break
            }

            case 'top10': {
              if (typeof val === 'number') {
                if (rule.bottom) matches = val <= bottomCutoff
                else matches = val >= topCutoff
              }
              break
            }

            case 'aboveAverage': {
              if (typeof val === 'number') {
                if (rule.above !== false) {
                  matches = rule.equalOrAbove ? val >= avg : val > avg
                } else {
                  matches = rule.equalOrAbove ? val <= avg : val < avg
                }
              }
              break
            }

            case 'expression': {
              if (rule.formula) {
                const dr = r - baseR
                const dc = c - baseC
                const shifted = shiftFormula(rule.formula, dr, dc)
                try {
                  const evalRes = store.hf.calculateFormula(shifted, store.hfId(sheet))
                  matches = evalRes === true || (typeof evalRes === 'number' && evalRes !== 0)
                } catch {
                  matches = false
                }
              }
              break
            }

            case 'dataBar': {
              if (typeof val === 'number' && rule.dataBar) {
                const range = maxVal - minVal || 1
                const fraction = Math.max(0, Math.min(1, (val - minVal) / range))
                dataBar = { fraction, color: rule.dataBar.color, negative: val < 0 }
                matches = true
              }
              break
            }

            case 'colorScale': {
              if (typeof val === 'number' && rule.colorScale) {
                const range = maxVal - minVal || 1
                const t = Math.max(0, Math.min(1, (val - minVal) / range))
                if (rule.colorScale.midColor) {
                  colorScale = t < 0.5
                    ? interpolateColor(rule.colorScale.minColor, rule.colorScale.midColor, t * 2)
                    : interpolateColor(rule.colorScale.midColor, rule.colorScale.maxColor, (t - 0.5) * 2)
                } else {
                  colorScale = interpolateColor(rule.colorScale.minColor, rule.colorScale.maxColor, t)
                }
                matches = true
              }
              break
            }

            case 'iconSet': {
              if (typeof val === 'number' && rule.iconSet) {
                const range = maxVal - minVal || 1
                const t = Math.max(0, Math.min(1, (val - minVal) / range))
                const count = rule.iconSet.name.startsWith('4') ? 4 : rule.iconSet.name.startsWith('5') ? 5 : 3
                let idx = Math.min(count - 1, Math.floor(t * count))
                if (rule.iconSet.reverse) idx = count - 1 - idx
                icon = { set: rule.iconSet.name, index: idx }
                matches = true
              }
              break
            }
          }

          if (matches) {
            const next: CfEvaluation = cur ? { ...cur } : {}
            if (rule.format) {
              next.style = { ...(next.style || {}), ...rule.format }
            }
            if (dataBar) next.dataBar = dataBar
            if (colorScale) next.colorScale = colorScale
            if (icon) next.icon = icon
            if (rule.stopIfTrue) (next as any)._stopped = true
            result.set(key, next)
          }
        }
      }
    }
  }

  return result
}

/** Evaluates conditional formatting for a single cell, using cached sheet evaluations. */
export function evaluateCf(
  store: WorkbookStore,
  sheet: SheetState,
  r: number,
  c: number
): CfEvaluation {
  const cached = sheetCache.get(sheet.uid)
  if (!cached || cached.version !== store.version) {
    const evaluations = buildSheetEvaluations(store, sheet)
    sheetCache.set(sheet.uid, { version: store.version, evaluations })
    return evaluations.get(`${r},${c}`) || {}
  }
  return cached.evaluations.get(`${r},${c}`) || {}
}

/** Adds a new rule to the sheet and triggers undoable updateLayout. */
export function addCfRule(store: WorkbookStore, sheet: SheetState, rule: Omit<CfRule, 'id' | 'priority'>): void {
  store.updateLayout(sheet, 'Conditional Formatting', (l) => {
    const list = [...(l.conditionalFormats || [])]
    const priority = list.length + 1
    const id = 'cf_' + Math.random().toString(36).slice(2, 9)
    list.push({ ...rule, id, priority })
    l.conditionalFormats = list
  })
}

/** Clears conditional formatting rules from selection or whole sheet. */
export function clearCfRules(store: WorkbookStore, sheet: SheetState, scope: 'selection' | 'sheet'): void {
  store.updateLayout(sheet, 'Clear Conditional Formatting', (l) => {
    if (scope === 'sheet') {
      l.conditionalFormats = []
    } else {
      const selRanges = sheet.selection.ranges
      const remaining: CfRule[] = []
      for (const rule of l.conditionalFormats || []) {
        const nextRanges = rule.ranges.filter((rg) => !selRanges.some((sel) => rangesOverlap(rg, sel)))
        if (nextRanges.length > 0) {
          remaining.push({ ...rule, ranges: nextRanges })
        }
      }
      l.conditionalFormats = remaining
    }
  })
}

function rangesOverlap(a: Range, b: Range): boolean {
  const na = normRange(a)
  const nb = normRange(b)
  return !(na.r2 < nb.r1 || na.r1 > nb.r2 || na.c2 < nb.c1 || na.c1 > nb.c2)
}

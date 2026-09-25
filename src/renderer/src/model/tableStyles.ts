// Excel Table Styles (Home > Format as Table / Table Design > Table Styles gallery): the
// Light1-21 / Medium1-28 / Dark1-11 palette, derived from the Office theme accent colors the same
// way real Excel's built-in table styles do - each style rotates through accent1..accent6 and
// varies tint/shade and which bands/borders are colored. This is an approximation of Excel's exact
// built-in style definitions (which aren't published as a formula), tuned to look right rather than
// byte-match Excel's theme XML; see the module doc in the PR/report for what's approximated.
import { THEME_COLORS } from './types'

export type TableStyleClass = 'light' | 'medium' | 'dark'

export interface TableStyleColors {
  name: string
  cls: TableStyleClass
  /** Header row: undefined fill = no fill (Light styles mostly just underline the header). */
  headerFill?: string
  headerFont: string
  headerBold: boolean
  /** Body banding fills. band2 is the "off" band; band1 undefined = no band color (odd/plain rows). */
  band1?: string
  band2?: string
  /** Outline/border color drawn around the whole table and between header/body. */
  border?: string
  totalFont: string
  totalBold: boolean
  /** First/Last column emphasis fill+font, used only when those options are checked. */
  firstColFill?: string
  firstColFont?: string
  lastColFill?: string
  lastColFont?: string
}

const ACCENTS = [THEME_COLORS.accent1, THEME_COLORS.accent2, THEME_COLORS.accent3, THEME_COLORS.accent4, THEME_COLORS.accent5, THEME_COLORS.accent6]

function tint(hex: string, t: number): string {
  const h = hex.replace('#', '')
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16)
  const adj = (c: number) => (t < 0 ? Math.round(c * (1 + t)) : Math.round(c * (1 - t) + 255 * t))
  const toHex = (n: number) => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')
  return '#' + toHex(adj(r)) + toHex(adj(g)) + toHex(adj(b))
}

function buildLight(i: number): TableStyleColors {
  // Light1 is the plain "no color, thin borders" style Excel puts first in the gallery.
  if (i === 1) {
    return {
      name: 'TableStyleLight1', cls: 'light', headerFont: '#000000', headerBold: true,
      border: '#000000', totalFont: '#000000', totalBold: true
    }
  }
  const accent = ACCENTS[(i - 2) % 6]
  return {
    name: 'TableStyleLight' + i, cls: 'light', headerFont: '#000000', headerBold: true,
    border: accent, band1: undefined, band2: tint(accent, 0.8),
    totalFont: '#000000', totalBold: true,
    firstColFill: undefined, firstColFont: '#000000', lastColFill: undefined, lastColFont: '#000000'
  }
}

function buildMedium(i: number): TableStyleColors {
  const accent = ACCENTS[(i - 1) % 6]
  const dark = (i - 1) % 12 >= 6 // second half of every 12 darkens the header instead of tinting
  return {
    name: 'TableStyleMedium' + i, cls: 'medium',
    headerFill: dark ? tint(accent, -0.25) : accent,
    headerFont: '#FFFFFF', headerBold: true,
    band1: undefined, band2: tint(accent, 0.8),
    border: tint(accent, -0.1),
    totalFont: '#000000', totalBold: true,
    firstColFill: tint(accent, -0.1), firstColFont: '#FFFFFF',
    lastColFill: tint(accent, -0.1), lastColFont: '#FFFFFF'
  }
}

function buildDark(i: number): TableStyleColors {
  const accent = ACCENTS[(i - 1) % 6]
  return {
    name: 'TableStyleDark' + i, cls: 'dark',
    headerFill: tint(accent, -0.35),
    headerFont: '#FFFFFF', headerBold: true,
    band1: tint(accent, -0.35), band2: tint(accent, -0.15),
    border: tint(accent, -0.45),
    totalFont: '#FFFFFF', totalBold: true,
    firstColFill: tint(accent, -0.45), firstColFont: '#FFFFFF',
    lastColFill: tint(accent, -0.45), lastColFont: '#FFFFFF'
  }
}

export const LIGHT_STYLES: TableStyleColors[] = Array.from({ length: 21 }, (_, k) => buildLight(k + 1))
export const MEDIUM_STYLES: TableStyleColors[] = Array.from({ length: 28 }, (_, k) => buildMedium(k + 1))
export const DARK_STYLES: TableStyleColors[] = Array.from({ length: 11 }, (_, k) => buildDark(k + 1))
export const ALL_TABLE_STYLES: TableStyleColors[] = [...LIGHT_STYLES, ...MEDIUM_STYLES, ...DARK_STYLES]
/** Excel's actual gallery default when you press Ctrl+T with no style already chosen. */
export const DEFAULT_TABLE_STYLE = 'TableStyleMedium2'

export function tableStyleByName(name: string): TableStyleColors | undefined {
  return ALL_TABLE_STYLES.find((s) => s.name === name)
}

// Excel 2021 (Office theme) default chart colors and typography, matching the built-in "Office" theme
// used by Office 2013 through 2021/LTSC (before the 2023+ "Aptos"/blue-accent refresh).
export const CHART_COLORS = [
  '#4472C4', // accent1 (blue)
  '#ED7D31', // accent2 (orange)
  '#A5A5A5', // accent3 (gray)
  '#FFC000', // accent4 (gold)
  '#5B9BD5', // accent5 (light blue)
  '#70AD47' // accent6 (green)
]

/** Alternate "Change Colors" gallery variants: monochromatic ramps of one accent color. */
export function colorSet(setId: number, count: number): string[] {
  if (setId <= 0) return Array.from({ length: count }, (_, i) => CHART_COLORS[i % CHART_COLORS.length])
  const base = CHART_COLORS[(setId - 1) % CHART_COLORS.length]
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(base.slice(i, i + 2), 16))
  return Array.from({ length: count }, (_, i) => {
    const t = count <= 1 ? 0 : i / (count - 1)
    const tint = 0.15 + t * 0.65 // lighten progressively
    const mix = (c: number): number => Math.round(c + (255 - c) * tint)
    return `#${[mix(r), mix(g), mix(b)].map((n) => n.toString(16).padStart(2, '0')).join('')}`
  })
}

export const CHART_FONT = 'Calibri, "Carlito", "Segoe UI", sans-serif'
export const GRIDLINE_COLOR = '#D9D9D9'
export const AXIS_COLOR = '#808080'
export const TITLE_COLOR = '#404040'
export const TEXT_COLOR = '#595959'

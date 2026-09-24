// Office 2023+ default chart theme: series colors and typography, matching Excel's built-in style.
export const CHART_COLORS = [
  '#156082', // accent1 (blue)
  '#E97132', // accent2 (orange)
  '#196B24', // accent3 (green)
  '#0F9ED5', // accent4 (light blue)
  '#A02B93', // accent5 (purple/magenta)
  '#4EA72E' // accent6 (light green)
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

export const CHART_FONT = 'Aptos, "Aptos Narrow", "Segoe UI", sans-serif'
export const GRIDLINE_COLOR = '#D9D9D9'
export const AXIS_COLOR = '#808080'
export const TITLE_COLOR = '#404040'
export const TEXT_COLOR = '#595959'

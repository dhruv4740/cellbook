// Page Layout > Themes / Colors / Fonts / Effects. A theme here is a colour palette, a font pair and an effects name.
// Colours are listed in the picker's column order: lt1, dk1, lt2, dk2, accent1..accent6.
export interface ThemeColors { name: string; colors: string[] }
export interface ThemeFonts { name: string; major: string; minor: string }
export interface ThemeChoice { colors?: ThemeColors; fonts?: ThemeFonts; effects?: string }

export const OFFICE_COLORS: ThemeColors = { name: 'Office', colors: ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'] }
export const COLOR_SCHEMES: ThemeColors[] = [
  OFFICE_COLORS,
  { name: 'Office 2007 - 2010', colors: ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'] },
  { name: 'Office 2023', colors: ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'] },
  { name: 'Blue', colors: ['FFFFFF', '000000', 'DBEFF9', '17406D', '0F6FC6', '009DD9', '0BD0D9', '10CF9B', '7CCA62', 'A5C249'] },
  { name: 'Green', colors: ['FFFFFF', '000000', 'E3DED1', '455F51', '549E39', '8AB833', 'C0CF3A', '029676', '4AB5C4', '0989B1'] }
]
export const FONT_SCHEMES: ThemeFonts[] = [
  { name: 'Office', major: 'Calibri Light', minor: 'Calibri' },
  { name: 'Office 2007 - 2010', major: 'Cambria', minor: 'Calibri' },
  { name: 'Arial', major: 'Arial', minor: 'Arial' },
  { name: 'Georgia', major: 'Georgia', minor: 'Georgia' },
  { name: 'Century Gothic', major: 'Century Gothic', minor: 'Century Gothic' },
  { name: 'Times New Roman', major: 'Times New Roman', minor: 'Times New Roman' }
]
/** Effects have no rendering in a grid (they style shapes' bevel/shadow in Excel); only the name is kept. */
export const EFFECT_SCHEMES = ['Office', 'Subtle', 'Moderate', 'Intense']
export const THEMES: { name: string; colors: ThemeColors; fonts: ThemeFonts }[] = [
  { name: 'Office', colors: OFFICE_COLORS, fonts: FONT_SCHEMES[0] },
  { name: 'Office 2007 - 2010', colors: COLOR_SCHEMES[1], fonts: FONT_SCHEMES[1] },
  { name: 'Blue', colors: COLOR_SCHEMES[3], fonts: FONT_SCHEMES[2] },
  { name: 'Green', colors: COLOR_SCHEMES[4], fonts: FONT_SCHEMES[3] }
]

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

/** Rewrites the colour scheme and the major/minor latin fonts of a theme1.xml, leaving everything else intact. */
export function applyThemeToXml(xml: string, t: ThemeChoice): string {
  let out = xml
  if (t.colors) {
    const c = t.colors.colors
    const slots: [string, string][] = [['lt1', c[0]], ['dk1', c[1]], ['lt2', c[2]], ['dk2', c[3]], ['accent1', c[4]], ['accent2', c[5]], ['accent3', c[6]], ['accent4', c[7]], ['accent5', c[8]], ['accent6', c[9]]]
    for (const [slot, hex] of slots) {
      out = out.replace(new RegExp(`(<a:${slot}>)[\\s\\S]*?(</a:${slot}>)`), `$1<a:srgbClr val="${hex}"/>$2`)
    }
    out = out.replace(/(<a:clrScheme\b[^>]*\bname=")[^"]*(")/, `$1${esc(t.colors.name)}$2`)
  }
  if (t.fonts) {
    out = out.replace(/(<a:majorFont>\s*<a:latin\b[^>]*\btypeface=")[^"]*(")/, `$1${esc(t.fonts.major)}$2`)
    out = out.replace(/(<a:minorFont>\s*<a:latin\b[^>]*\btypeface=")[^"]*(")/, `$1${esc(t.fonts.minor)}$2`)
    out = out.replace(/(<a:fontScheme\b[^>]*\bname=")[^"]*(")/, `$1${esc(t.fonts.name)}$2`)
  }
  return out
}

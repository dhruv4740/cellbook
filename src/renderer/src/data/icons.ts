// Insert > Icons: search the curated set, rasterize the chosen icon's SVG to a PNG and insert it as a picture.
import { ICON_CATEGORIES, IconEntry } from './iconSet'
import { insertPictureFromDataUrl } from './pictures'

export const ALL_ICONS: (IconEntry & { category: string })[] = ICON_CATEGORIES.flatMap((c) => c.icons.map((i) => ({ ...i, category: c.name })))

/** Icons of `category` (or every icon for null) whose name or keywords contain each word of `query`. */
export function searchIcons(query: string, category: string | null): (IconEntry & { category: string })[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean)
  return ALL_ICONS.filter((i) => (!category || i.category === category) && words.every((w) => i.keywords.includes(w)))
}

/** Turns a rendered Fluent icon (1em, currentColor) into a standalone SVG document of `px` x `px` in `color`. */
export function standaloneSvg(markup: string, color: string, px: number): string {
  let svg = markup.replace(/\sclass="[^"]*"/, '').replace(/\sstyle="[^"]*"/, '').replace(/currentColor/g, color)
  svg = svg.replace(/\swidth="[^"]*"/, ` width="${px}"`).replace(/\sheight="[^"]*"/, ` height="${px}"`)
  if (!/xmlns=/.test(svg)) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"')
  return svg
}

/** Draws the icon on a canvas and inserts the PNG at the active cell (Excel inserts icons as editable graphics; this is a bitmap). */
export async function insertIcon(markup: string, color: string, label: string): Promise<void> {
  const px = 256
  const img = new Image()
  img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(standaloneSvg(markup, color, px))
  await img.decode()
  const canvas = document.createElement('canvas')
  canvas.width = px; canvas.height = px
  canvas.getContext('2d')!.drawImage(img, 0, 0, px, px)
  await insertPictureFromDataUrl(canvas.toDataURL('image/png'), label)
}

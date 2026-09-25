// Page geometry: inches -> CSS px (96dpi, matching Excel's screen-to-paper mapping) and the
// printable-area math shared by the pagination engine and the page renderer.
import { PageSetup, PAPER_SIZES } from '../model/types'

export const PX_PER_INCH = 96

export function inToPx(inches: number): number {
  return inches * PX_PER_INCH
}

export interface PageGeometry {
  /** Paper size in px, already oriented (landscape swaps w/h). */
  paperWidthPx: number
  paperHeightPx: number
  marginTopPx: number
  marginBottomPx: number
  marginLeftPx: number
  marginRightPx: number
  headerPx: number
  footerPx: number
  /** Area available for the grid, i.e. paper minus margins (header/footer sit inside top/bottom margin). */
  printableWidthPx: number
  printableHeightPx: number
}

export function pageGeometry(ps: PageSetup): PageGeometry {
  const paper = PAPER_SIZES[ps.paperSize] ?? PAPER_SIZES.letter
  const wIn = ps.orientation === 'landscape' ? paper.h : paper.w
  const hIn = ps.orientation === 'landscape' ? paper.w : paper.h
  const m = ps.margins
  return {
    paperWidthPx: inToPx(wIn),
    paperHeightPx: inToPx(hIn),
    marginTopPx: inToPx(m.top),
    marginBottomPx: inToPx(m.bottom),
    marginLeftPx: inToPx(m.left),
    marginRightPx: inToPx(m.right),
    headerPx: inToPx(m.header),
    footerPx: inToPx(m.footer),
    printableWidthPx: inToPx(wIn - m.left - m.right),
    printableHeightPx: inToPx(hIn - m.top - m.bottom)
  }
}

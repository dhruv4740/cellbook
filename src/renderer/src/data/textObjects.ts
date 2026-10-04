// Insert > Text > WordArt / Signature Line and Insert > Symbols > Equation: text boxes pre-filled and pre-styled.
// All three are ordinary text-box shapes, so move / resize / arrange / edit-text and the xlsx drawing round trip come for free.
import { ShapeDef } from '../model/types'
import { insertShape } from './shapes'

export interface WordArtStyle { label: string; patch: Partial<ShapeDef> }

const wa = (label: string, color: string, extra: Partial<ShapeDef> = {}): WordArtStyle => ({
  label,
  patch: { text: 'Your text here', fontSize: 32, bold: true, color, align: 'ctr', width: 280, height: 64, fontFamily: 'Calibri', ...extra }
})

/** Excel's WordArt gallery, reduced to what a text box can draw: fill colour, outline colour and a drop shadow. */
export const WORDART_STYLES: WordArtStyle[] = [
  wa('Fill: Black, Text 1; Shadow', '#000000', { textShadow: true }),
  wa('Fill: Blue, Accent 1; Shadow', '#4472c4', { textShadow: true }),
  wa('Fill: Orange, Accent 2; Outline: Orange, Accent 2', '#ed7d31', { textOutline: '#c55a11' }),
  wa('Fill: Gray, Accent 3; Shadow', '#a5a5a5', { textShadow: true }),
  wa('Fill: Gold, Accent 4; Outline: Gold, Accent 4', '#ffc000', { textOutline: '#bf9000' }),
  wa('Fill: Blue, Accent 5; Shadow', '#5b9bd5', { textShadow: true }),
  wa('Fill: Green, Accent 6; Shadow', '#70ad47', { textShadow: true }),
  wa('Fill: White; Outline: Blue, Accent 1; Shadow', '#ffffff', { textOutline: '#4472c4', textShadow: true }),
  wa('Fill: White; Outline: Orange, Accent 2; Shadow', '#ffffff', { textOutline: '#ed7d31', textShadow: true })
]

export function insertWordArt(style: WordArtStyle): void {
  insertShape('textbox', style.patch,'Insert WordArt')
}

/** Excel's signature line: an X, the line to sign on, and the signer's name beneath. */
export const SIGNATURE_TEXT = (signer = 'Signer'): string => `X\n${'_'.repeat(32)}\n${signer}`

export function insertSignatureLine(): void {
  insertShape('textbox', { text: SIGNATURE_TEXT(), fontSize: 12, width: 230, height: 70, fontFamily: 'Calibri' }, 'Insert Signature Line')
}

/** The built-in equation gallery as Unicode text (a real OMML equation editor is not implemented). */
export const EQUATIONS: { label: string; text: string }[] = [
  { label: 'Area of Circle', text: 'A = πr²' },
  { label: 'Binomial Theorem', text: '(x + a)ⁿ = ∑ₖ₌₀ⁿ (ⁿₖ) xᵏ aⁿ⁻ᵏ' },
  { label: 'Expansion of a Sum', text: '(1 + x)ⁿ = 1 + nx/1! + n(n − 1)x²/2! + ⋯' },
  { label: 'Fourier Series', text: 'f(x) = a₀ + ∑ₙ₌₁^∞ (aₙ cos(nπx/L) + bₙ sin(nπx/L))' },
  { label: 'Pythagorean Theorem', text: 'a² + b² = c²' },
  { label: 'Quadratic Formula', text: 'x = (−b ± √(b² − 4ac)) / 2a' },
  { label: 'Taylor Expansion', text: 'eˣ = 1 + x/1! + x²/2! + x³/3! + ⋯, −∞ < x < ∞' },
  { label: 'Trig Identity 1', text: 'sin α ± sin β = 2 sin ½(α ± β) cos ½(α ∓ β)' },
  { label: 'Trig Identity 2', text: 'cos α + cos β = 2 cos ½(α + β) cos ½(α − β)' }
]
export const NEW_EQUATION_TEXT = 'Type equation here.'

export function insertEquation(text: string): void {
  insertShape('textbox', { text, fontSize: 16, italic: true, fontFamily: 'Cambria Math', width: Math.max(150, Math.min(420, text.length * 11 + 20)), height: 48 }, 'Insert Equation')
}

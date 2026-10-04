import { CURRENCY_SYMBOL, DAY_FIRST } from '../engine/locale'
import { useState, useRef, useEffect, MouseEvent as ReactMouseEvent } from 'react'
import { Dialog } from '../components/Dialog'
import {
  BorderLineStyle,
  BorderSide,
  CellStyle,
  CellValue,
  normRange
} from '../model/types'
import {
  FORMAT_CATEGORIES,
  buildNumberFormat,
  formatValue
} from '../format/numfmt'
import { store, useStore } from '../model/store'
import { OFFICE_COLORS } from '../data/themes'
import { mergeCells } from '../grid/commands'
import './FormatCellsDialog.css'

export type FormatCellsTab =
  | 'Number'
  | 'Alignment'
  | 'Font'
  | 'Border'
  | 'Fill'
  | 'Protection'

export interface FormatCellsDialogProps {
  initialTab?: string
  initialStyle?: CellStyle
  cellValue?: CellValue
  onApply: (patch: Partial<CellStyle>) => void
  onClose: () => void
}

const TABS: FormatCellsTab[] = ['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection']

function shade(hex: string, amt: number): string {
  const n = parseInt(hex.slice(1), 16)
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255
  const mix = (c: number) => (amt > 0 ? c + (255 - c) * amt : c * (1 + amt))
  r = Math.round(mix(r)); g = Math.round(mix(g)); b = Math.round(mix(b))
  return '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase()
}
const themeRows = (): string[][] => {
  const base = (store.theme.colors ?? OFFICE_COLORS).colors.map((c) => '#' + c)
  return [0, 0.8, 0.6, 0.4, -0.25, -0.5].map((amt, i) => (i === 0 ? base : base.map((c) => shade(c, amt))))
}
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0']

const FONT_FAMILIES = [
  'Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Comic Sans MS',
  'Consolas', 'Courier New', 'Georgia', 'Impact', 'Segoe UI', 'Tahoma',
  'Times New Roman', 'Trebuchet MS', 'Verdana'
]

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

// 14 Excel Line Styles in 2 columns
const LINE_STYLES: { id: number; label: string; style: BorderLineStyle | null; borderCss: string }[] = [
  // Col 1
  { id: 0, label: 'None', style: null, borderCss: 'none' },
  { id: 1, label: 'Thin', style: 'thin', borderCss: '1px solid #000' },
  { id: 2, label: 'Medium dashed', style: 'mediumDashed', borderCss: '2px dashed #000' },
  { id: 3, label: 'Dashed', style: 'dashed', borderCss: '1px dashed #000' },
  { id: 4, label: 'Dotted', style: 'dotted', borderCss: '1px dotted #000' },
  { id: 5, label: 'Thick', style: 'thick', borderCss: '3px solid #000' },
  { id: 6, label: 'Double', style: 'double', borderCss: '3px double #000' },
  // Col 2
  { id: 7, label: 'Hair', style: 'hair', borderCss: '1px dotted #777' },
  { id: 8, label: 'Medium', style: 'medium', borderCss: '2px solid #000' },
  { id: 9, label: 'Dash dot', style: 'dashDot', borderCss: '1px dashed #000' },
  { id: 10, label: 'Medium dash dot', style: 'mediumDashDot', borderCss: '2px dashed #000' },
  { id: 11, label: 'Dash dot dot', style: 'dashDotDot', borderCss: '1px dotted #000' },
  { id: 12, label: 'Medium dash dot dot', style: 'mediumDashDotDot', borderCss: '2px dashed #000' },
  { id: 13, label: 'Slant dash dot', style: 'slantDashDot', borderCss: '2px dashed #000' }
]

export function FormatCellsDialog({
  initialTab = 'Number',
  initialStyle = {},
  cellValue,
  onApply,
  onClose
}: FormatCellsDialogProps) {
  const store = useStore()
  const sheet = store.active
  const ranges = sheet.selection.ranges
  const isMultiCell = ranges.length > 1 || (
    ranges[0] && (ranges[0].r1 !== ranges[0].r2 || ranges[0].c1 !== ranges[0].c2)
  )

  // Check if active cell is in a merged range
  const activeCell = sheet.selection.active
  const initialMerged = !!store.mergeAt(sheet, activeCell.r, activeCell.c)

  const [tab, setTab] = useState<FormatCellsTab>(
    (TABS.find((t) => t.toLowerCase() === initialTab.toLowerCase()) as FormatCellsTab) || 'Number'
  )

  // Track modified fields
  const modified = useRef<Set<string>>(new Set())

  // --- Number State ---
  const [numCategory, setNumCategory] = useState<string>(() => {
    const f = initialStyle.numFmt
    if (!f || f === 'General') return 'General'
    for (const [cat, entries] of Object.entries(FORMAT_CATEGORIES)) {
      if (entries.some((e) => e.code === f)) return cat
    }
    return 'Custom'
  })
  const [numCode, setNumCode] = useState<string>(initialStyle.numFmt || 'General')
  const [numDecimals, setNumDecimals] = useState<number>(2)
  const [numThousands, setNumThousands] = useState<boolean>(false)
  const [numNegative, setNumNegative] = useState<0 | 1 | 2 | 3>(0)
  const [currencySymbol, setCurrencySymbol] = useState<string>(CURRENCY_SYMBOL.length === 1 ? CURRENCY_SYMBOL : '$')

  // --- Alignment State ---
  const [hAlign, setHAlign] = useState<CellStyle['hAlign']>(initialStyle.hAlign || 'general')
  const [vAlign, setVAlign] = useState<CellStyle['vAlign']>(initialStyle.vAlign || 'bottom')
  const [indent, setIndent] = useState<number>(initialStyle.indent || 0)
  const [justifyDistributed, setJustifyDistributed] = useState<boolean>(false)
  const [wrap, setWrap] = useState<boolean>(!!initialStyle.wrap)
  const [shrinkToFit, setShrinkToFit] = useState<boolean>(!!initialStyle.shrinkToFit)
  const [merged, setMerged] = useState<boolean>(initialMerged)
  const [rotation, setRotation] = useState<number>(initialStyle.rotation ?? 0)

  // --- Font State ---
  const [fontFamily, setFontFamily] = useState<string>(initialStyle.fontFamily || 'Aptos Narrow')
  const [fontFamilyInput, setFontFamilyInput] = useState<string>(initialStyle.fontFamily || 'Aptos Narrow')
  const [fontSize, setFontSize] = useState<number>(initialStyle.fontSize || 11)
  const [fontSizeInput, setFontSizeInput] = useState<string>(String(initialStyle.fontSize || 11))
  const [bold, setBold] = useState<boolean>(!!initialStyle.bold)
  const [italic, setItalic] = useState<boolean>(!!initialStyle.italic)
  const [fontStyleInput, setFontStyleInput] = useState<string>(
    initialStyle.bold && initialStyle.italic
      ? 'Bold Italic'
      : initialStyle.bold
        ? 'Bold'
        : initialStyle.italic
          ? 'Italic'
          : 'Regular'
  )
  const [underline, setUnderline] = useState<string>(initialStyle.underline || 'none')
  const [strike, setStrike] = useState<boolean>(!!initialStyle.strike)
  const [superscript, setSuperscript] = useState<boolean>(false)
  const [subscript, setSubscript] = useState<boolean>(false)
  const [fontColor, setFontColor] = useState<string | undefined>(initialStyle.color)
  const [fontColorOpen, setFontColorOpen] = useState(false)

  // --- Border State ---
  const [selectedStyleId, setSelectedStyleId] = useState<number>(1) // Thin
  const [borderColor, setBorderColor] = useState<string>('#000000')

  // Borders map: top, bottom, left, right, insideH, insideV, diagUp, diagDown
  const [borders, setBorders] = useState<{
    top?: BorderSide
    bottom?: BorderSide
    left?: BorderSide
    right?: BorderSide
    insideH?: BorderSide
    insideV?: BorderSide
    diagUp?: BorderSide
    diagDown?: BorderSide
  }>({
    top: initialStyle.borders?.top,
    bottom: initialStyle.borders?.bottom,
    left: initialStyle.borders?.left,
    right: initialStyle.borders?.right
  })

  // --- Fill State ---
  const [fill, setFill] = useState<string | undefined>(initialStyle.fill)
  const [patternColor, setPatternColor] = useState<string>('#000000')
  const [patternStyle, setPatternStyle] = useState<string>('Automatic')

  // --- Protection State ---
  const [locked, setLocked] = useState<boolean>(initialStyle.locked !== false)
  const [hidden, setHidden] = useState<boolean>(!!initialStyle.hidden)

  // Orientation drag handler
  const dialRef = useRef<HTMLCanvasElement>(null)
  const isDraggingDial = useRef(false)

  const drawDial = (angle: number) => {
    const canvas = dialRef.current
    if (!canvas) return
    const ctx = canvas.getContext('2d')
    if (!ctx) return

    const w = canvas.width, h = canvas.height
    ctx.clearRect(0, 0, w, h)

    const cx = 20, cy = h / 2
    const radius = 65

    // Arc
    ctx.beginPath()
    ctx.arc(cx, cy, radius, -Math.PI / 2, Math.PI / 2, false)
    ctx.strokeStyle = '#adadad'
    ctx.lineWidth = 1
    ctx.stroke()

    // Dots every 15 deg
    for (let deg = -90; deg <= 90; deg += 15) {
      const rad = (-deg * Math.PI) / 180
      const x = cx + radius * Math.cos(rad)
      const y = cy + radius * Math.sin(rad)
      ctx.beginPath()
      ctx.arc(x, y, deg % 45 === 0 ? 2.5 : 1.5, 0, 2 * Math.PI)
      ctx.fillStyle = '#666'
      ctx.fill()
    }

    // Pointer line & text
    const activeDeg = angle === 255 ? 0 : angle
    const rad = (-activeDeg * Math.PI) / 180
    const tipX = cx + radius * Math.cos(rad)
    const tipY = cy + radius * Math.sin(rad)

    ctx.beginPath()
    ctx.moveTo(cx, cy)
    ctx.lineTo(tipX, tipY)
    ctx.strokeStyle = '#000'
    ctx.lineWidth = 1.5
    ctx.stroke()

    // Red Diamond Handle
    ctx.beginPath()
    const dSize = 5
    ctx.moveTo(tipX, tipY - dSize)
    ctx.lineTo(tipX + dSize, tipY)
    ctx.lineTo(tipX, tipY + dSize)
    ctx.lineTo(tipX - dSize, tipY)
    ctx.closePath()
    ctx.fillStyle = '#e81123'
    ctx.fill()
    ctx.strokeStyle = '#a80000'
    ctx.lineWidth = 1
    ctx.stroke()

    // "Text" along current angle
    ctx.save()
    ctx.translate(cx + (radius * 0.45) * Math.cos(rad), cy + (radius * 0.45) * Math.sin(rad))
    ctx.rotate(rad)
    ctx.font = '11px Segoe UI, sans-serif'
    ctx.fillStyle = '#000'
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.fillText('Text', 0, -8)
    ctx.restore()
  }

  useEffect(() => {
    drawDial(rotation)
  }, [rotation])

  const updateAngleFromMouse = (e: MouseEvent | ReactMouseEvent) => {
    const canvas = dialRef.current
    if (!canvas) return
    const rect = canvas.getBoundingClientRect()
    const cx = rect.left + 20
    const cy = rect.top + canvas.height / 2
    const dx = e.clientX - cx
    const dy = e.clientY - cy
    let angleRad = -Math.atan2(dy, dx)
    let deg = Math.round((angleRad * 180) / Math.PI)
    if (deg > 90) deg = 90
    if (deg < -90) deg = -90
    setRotation(deg)
    modified.current.add('rotation')
  }

  const handleDialMouseDown = (e: ReactMouseEvent) => {
    isDraggingDial.current = true
    updateAngleFromMouse(e)
  }

  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      if (isDraggingDial.current) updateAngleFromMouse(e)
    }
    const onUp = () => {
      isDraggingDial.current = false
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [])

  // Number Category change
  const handleCategorySelect = (cat: string) => {
    setNumCategory(cat)
    modified.current.add('numFmt')
    if (cat === 'General') {
      setNumCode('General')
    } else if (cat === 'Number') {
      setNumCode(buildNumberFormat({ decimals: numDecimals, thousands: numThousands, negative: numNegative }))
    } else if (cat === 'Currency') {
      setNumCode(buildNumberFormat({ decimals: numDecimals, thousands: true, negative: numNegative, symbol: currencySymbol }))
    } else if (cat === 'Accounting') {
      setNumCode(currencySymbol ? `_(${currencySymbol}* #,##0.00_);_(${currencySymbol}* (#,##0.00);_(${currencySymbol}* "-"??_);_(@_)` : `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`)
    } else if (FORMAT_CATEGORIES[cat]?.[0]) {
      setNumCode(FORMAT_CATEGORIES[cat][0].code)
    }
  }

  // Border toggle
  const getSelectedBorderSide = (): BorderSide | undefined => {
    const item = LINE_STYLES.find((s) => s.id === selectedStyleId)
    if (!item || !item.style) return undefined
    return { style: item.style, color: borderColor }
  }

  const toggleBorder = (side: keyof typeof borders) => {
    modified.current.add('borders')
    setBorders((prev) => {
      const activeSide = getSelectedBorderSide()
      if (prev[side]) {
        const next = { ...prev }
        delete next[side]
        return next
      }
      if (!activeSide) return prev
      return { ...prev, [side]: activeSide }
    })
  }

  const handleBorderPreset = (preset: 'none' | 'outline' | 'inside') => {
    modified.current.add('borders')
    const activeSide = getSelectedBorderSide() || { style: 'thin', color: borderColor }
    if (preset === 'none') {
      setBorders({})
    } else if (preset === 'outline') {
      setBorders((prev) => ({
        ...prev,
        top: activeSide,
        bottom: activeSide,
        left: activeSide,
        right: activeSide
      }))
    } else if (preset === 'inside') {
      setBorders((prev) => ({
        ...prev,
        insideH: activeSide,
        insideV: activeSide
      }))
    }
  }

  const sampleDisplay = formatValue(cellValue ?? 1234.1, numCode === 'General' ? undefined : numCode).text

  const handleOk = () => {
    store.transact('Format Cells', () => {
      const patch: Partial<CellStyle> = {}

      if (modified.current.has('numFmt')) patch.numFmt = numCode
      if (modified.current.has('hAlign')) patch.hAlign = hAlign
      if (modified.current.has('vAlign')) patch.vAlign = vAlign
      if (modified.current.has('indent')) patch.indent = indent
      if (modified.current.has('wrap')) patch.wrap = wrap
      if (modified.current.has('shrinkToFit')) patch.shrinkToFit = shrinkToFit
      if (modified.current.has('rotation')) patch.rotation = rotation
      if (modified.current.has('fontFamily')) patch.fontFamily = fontFamily
      if (modified.current.has('fontSize')) patch.fontSize = fontSize
      if (modified.current.has('bold')) patch.bold = bold
      if (modified.current.has('italic')) patch.italic = italic
      if (modified.current.has('underline')) patch.underline = underline === 'none' ? undefined : (underline as any)
      if (modified.current.has('strike')) patch.strike = strike
      if (modified.current.has('color')) patch.color = fontColor
      if (modified.current.has('fill')) patch.fill = fill
      if (modified.current.has('locked')) patch.locked = locked
      if (modified.current.has('hidden')) patch.hidden = hidden

      // Apply borders
      if (modified.current.has('borders')) {
        // If inside borders were modified on a multi-cell selection
        if (isMultiCell && (borders.insideH || borders.insideV)) {
          // Bounding box of selection
          let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity
          ranges.forEach((rg) => {
            const n = normRange(rg)
            minR = Math.min(minR, n.r1)
            maxR = Math.max(maxR, n.r2)
            minC = Math.min(minC, n.c1)
            maxC = Math.max(maxC, n.c2)
          })

          store.applyStyle(sheet, ranges, (s, r, c) => {
            const curBorders = { ...s.borders }
            if (borders.top && r === minR) curBorders.top = borders.top
            if (borders.bottom && r === maxR) curBorders.bottom = borders.bottom
            if (borders.left && c === minC) curBorders.left = borders.left
            if (borders.right && c === maxC) curBorders.right = borders.right
            // Inside horizontal: top border on cells below the top row
            if (borders.insideH && r > minR) curBorders.top = borders.insideH
            // Inside vertical: left border on cells right of the leftmost column
            if (borders.insideV && c > minC) curBorders.left = borders.insideV
            return { ...s, ...patch, borders: curBorders }
          })
        } else {
          patch.borders = {
            top: borders.top,
            bottom: borders.bottom,
            left: borders.left,
            right: borders.right
          }
          store.applyStyle(sheet, ranges, patch)
        }
      } else if (Object.keys(patch).length > 0) {
        store.applyStyle(sheet, ranges, patch)
      }

      // Handle Merge Cells toggle
      if (modified.current.has('merged')) {
        if (merged && !initialMerged) {
          mergeCells('center')
        } else if (!merged && initialMerged) {
          mergeCells('unmerge')
        }
      }
    })

    onClose()
  }

  return (
    <Dialog title="Format Cells" width={620} onOk={handleOk} onCancel={onClose}>
      <div className="xl-fc-dialog">
        <div className="xl-fc-tabs">
          {TABS.map((t) => (
            <div
              key={t}
              className={`xl-fc-tab ${tab === t ? 'active' : ''}`}
              onClick={() => setTab(t)}
            >
              {t}
            </div>
          ))}
        </div>

        <div className="xl-fc-pane">
          {/* TAB 1: NUMBER */}
          {tab === 'Number' && (
            <div className="xl-fc-num-layout">
              <div className="xl-fc-num-left">
                <div style={{ marginBottom: 4, fontWeight: 500 }}>Category:</div>
                <div className="xl-fc-listbox" style={{ height: 350 }}>
                  {Object.keys(FORMAT_CATEGORIES).map((cat) => (
                    <div
                      key={cat}
                      className={`xl-fc-list-item ${numCategory === cat ? 'selected' : ''}`}
                      onClick={() => handleCategorySelect(cat)}
                    >
                      {cat}
                    </div>
                  ))}
                </div>
              </div>

              <div className="xl-fc-num-right">
                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Sample</span>
                  <div className="xl-fc-sample-box">{sampleDisplay}</div>
                </div>

                {numCategory === 'General' && (
                  <div className="xl-fc-help-text" style={{ marginTop: 12 }}>
                    General format cells have no specific number format.
                  </div>
                )}

                {numCategory === 'Number' && (
                  <div className="xl-fc-col">
                    <div className="xl-fc-row">
                      <span>Decimal places:</span>
                      <input
                        type="number"
                        min="0"
                        max="30"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={numDecimals}
                        onChange={(e) => {
                          const d = Math.max(0, +e.target.value)
                          setNumDecimals(d)
                          setNumCode(buildNumberFormat({ decimals: d, thousands: numThousands, negative: numNegative }))
                          modified.current.add('numFmt')
                        }}
                      />
                    </div>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={numThousands}
                        onChange={(e) => {
                          setNumThousands(e.target.checked)
                          setNumCode(buildNumberFormat({ decimals: numDecimals, thousands: e.target.checked, negative: numNegative }))
                          modified.current.add('numFmt')
                        }}
                      />
                      <span>Use 1000 Separator (,)</span>
                    </label>
                    <div style={{ marginTop: 8 }}>
                      <div style={{ marginBottom: 4 }}>Negative numbers:</div>
                      <div className="xl-fc-listbox" style={{ height: 110 }}>
                        {[
                          { val: 0, text: '-1234.10', color: '#000' },
                          { val: 1, text: '1234.10', color: 'red' },
                          { val: 2, text: '-1234.10', color: 'red' },
                          { val: 3, text: '(1234.10)', color: 'red' }
                        ].map((opt) => (
                          <div
                            key={opt.val}
                            className={`xl-fc-list-item ${numNegative === opt.val ? 'selected' : ''}`}
                            style={{ color: numNegative === opt.val ? '#fff' : opt.color }}
                            onClick={() => {
                              setNumNegative(opt.val as any)
                              setNumCode(buildNumberFormat({ decimals: numDecimals, thousands: numThousands, negative: opt.val as any }))
                              modified.current.add('numFmt')
                            }}
                          >
                            {opt.text}
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="xl-fc-help-text" style={{ marginTop: 12 }}>
                      Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary values.
                    </div>
                  </div>
                )}

                {numCategory === 'Currency' && (
                  <div className="xl-fc-col">
                    <div className="xl-fc-row">
                      <span>Decimal places:</span>
                      <input
                        type="number"
                        min="0"
                        max="30"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={numDecimals}
                        onChange={(e) => {
                          const d = Math.max(0, +e.target.value)
                          setNumDecimals(d)
                          setNumCode(buildNumberFormat({ decimals: d, thousands: true, negative: numNegative, symbol: currencySymbol }))
                          modified.current.add('numFmt')
                        }}
                      />
                    </div>
                    <div className="xl-fc-row">
                      <span>Symbol:</span>
                      <select
                        className="xl-fc-select"
                        value={currencySymbol}
                        onChange={(e) => {
                          setCurrencySymbol(e.target.value)
                          setNumCode(buildNumberFormat({ decimals: numDecimals, thousands: true, negative: numNegative, symbol: e.target.value === 'None' ? undefined : e.target.value }))
                          modified.current.add('numFmt')
                        }}
                      >
                        <option value="₹">₹</option>
                        <option value="$">$</option>
                        <option value="€">€</option>
                        <option value="£">£</option>
                        <option value="¥">¥</option>
                        <option value="None">None</option>
                      </select>
                    </div>
                    <div style={{ marginTop: 8 }}>
                      <div style={{ marginBottom: 4 }}>Negative numbers:</div>
                      <div className="xl-fc-listbox" style={{ height: 100 }}>
                        {[
                          { val: 0, text: `-${currencySymbol}1,234.10`, color: '#000' },
                          { val: 1, text: `${currencySymbol}1,234.10`, color: 'red' },
                          { val: 2, text: `-${currencySymbol}1,234.10`, color: 'red' },
                          { val: 3, text: `(${currencySymbol}1,234.10)`, color: 'red' }
                        ].map((opt) => (
                          <div
                            key={opt.val}
                            className={`xl-fc-list-item ${numNegative === opt.val ? 'selected' : ''}`}
                            style={{ color: numNegative === opt.val ? '#fff' : opt.color }}
                            onClick={() => {
                              setNumNegative(opt.val as any)
                              setNumCode(buildNumberFormat({ decimals: numDecimals, thousands: true, negative: opt.val as any, symbol: currencySymbol }))
                              modified.current.add('numFmt')
                            }}
                          >
                            {opt.text}
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="xl-fc-help-text" style={{ marginTop: 8 }}>
                      Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.
                    </div>
                  </div>
                )}

                {numCategory === 'Accounting' && (
                  <div className="xl-fc-col">
                    <div className="xl-fc-row">
                      <span>Decimal places:</span>
                      <input
                        type="number"
                        min="0"
                        max="30"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={numDecimals}
                        onChange={(e) => {
                          setNumDecimals(Math.max(0, +e.target.value))
                          modified.current.add('numFmt')
                        }}
                      />
                    </div>
                    <div className="xl-fc-row">
                      <span>Symbol:</span>
                      <select
                        className="xl-fc-select"
                        value={currencySymbol}
                        onChange={(e) => {
                          setCurrencySymbol(e.target.value)
                          const sym = e.target.value === 'None' ? '' : e.target.value
                          setNumCode(sym ? `_(${sym}* #,##0.00_);_(${sym}* (#,##0.00);_(${sym}* "-"??_);_(@_)` : `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`)
                          modified.current.add('numFmt')
                        }}
                      >
                        <option value="₹">₹</option>
                        <option value="$">$</option>
                        <option value="None">None</option>
                      </select>
                    </div>
                    <div className="xl-fc-help-text" style={{ marginTop: 8 }}>
                      Accounting formats align the currency symbols and decimal points in a column.
                    </div>
                  </div>
                )}

                {['Date', 'Time', 'Fraction', 'Special'].includes(numCategory) && (
                  <div className="xl-fc-col">
                    <div style={{ marginBottom: 4 }}>Type:</div>
                    <div className="xl-fc-listbox" style={{ height: 180 }}>
                      {FORMAT_CATEGORIES[numCategory]?.map((entry) => (
                        <div
                          key={entry.code}
                          className={`xl-fc-list-item ${numCode === entry.code ? 'selected' : ''}`}
                          onClick={() => {
                            setNumCode(entry.code)
                            modified.current.add('numFmt')
                          }}
                        >
                          {entry.label}
                        </div>
                      ))}
                    </div>
                    <div className="xl-fc-row" style={{ marginTop: 8 }}>
                      <span>Locale (location):</span>
                      <select className="xl-fc-select" style={{ width: 180 }} disabled>
                        <option>{DAY_FIRST ? 'English (India)' : 'English (United States)'}</option>
                      </select>
                    </div>
                  </div>
                )}

                {numCategory === 'Percentage' && (
                  <div className="xl-fc-col">
                    <div className="xl-fc-row">
                      <span>Decimal places:</span>
                      <input
                        type="number"
                        min="0"
                        max="30"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={numDecimals}
                        onChange={(e) => {
                          const d = Math.max(0, +e.target.value)
                          setNumDecimals(d)
                          setNumCode(d > 0 ? '0.' + '0'.repeat(d) + '%' : '0%')
                          modified.current.add('numFmt')
                        }}
                      />
                    </div>
                    <div className="xl-fc-help-text" style={{ marginTop: 8 }}>
                      Percentage formats multiply the cell value by 100 and display the result with a percent symbol.
                    </div>
                  </div>
                )}

                {numCategory === 'Scientific' && (
                  <div className="xl-fc-col">
                    <div className="xl-fc-row">
                      <span>Decimal places:</span>
                      <input
                        type="number"
                        min="0"
                        max="30"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={numDecimals}
                        onChange={(e) => {
                          const d = Math.max(0, +e.target.value)
                          setNumDecimals(d)
                          setNumCode(d > 0 ? '0.' + '0'.repeat(d) + 'E+00' : '0E+00')
                          modified.current.add('numFmt')
                        }}
                      />
                    </div>
                    <div className="xl-fc-help-text" style={{ marginTop: 8 }}>
                      Scientific formats display numbers in exponential notation.
                    </div>
                  </div>
                )}

                {numCategory === 'Text' && (
                  <div className="xl-fc-help-text" style={{ marginTop: 12 }}>
                    Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.
                  </div>
                )}

                {numCategory === 'Custom' && (
                  <div className="xl-fc-col">
                    <div className="xl-fc-row">
                      <span>Type:</span>
                      <input
                        className="xl-fc-input"
                        style={{ flex: 1 }}
                        value={numCode}
                        onChange={(e) => {
                          setNumCode(e.target.value)
                          modified.current.add('numFmt')
                        }}
                      />
                      <button className="xl-btn" style={{ minWidth: 60 }} disabled>
                        Delete
                      </button>
                    </div>
                    <div className="xl-fc-listbox" style={{ height: 180, marginTop: 4 }}>
                      {FORMAT_CATEGORIES.Custom.map((entry) => (
                        <div
                          key={entry.code}
                          className={`xl-fc-list-item ${numCode === entry.code ? 'selected' : ''}`}
                          onClick={() => {
                            setNumCode(entry.code)
                            modified.current.add('numFmt')
                          }}
                        >
                          {entry.code}
                        </div>
                      ))}
                    </div>
                    <div className="xl-fc-help-text" style={{ marginTop: 6 }}>
                      Type the number format code, using one of the existing codes as a starting point.
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 2: ALIGNMENT */}
          {tab === 'Alignment' && (
            <div className="xl-fc-align-layout">
              <div className="xl-fc-align-left">
                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Text alignment</span>
                  <div className="xl-fc-col">
                    <div className="xl-fc-row" style={{ justifyContent: 'space-between' }}>
                      <span>Horizontal:</span>
                      <select
                        className="xl-fc-select"
                        style={{ width: 170 }}
                        value={hAlign}
                        onChange={(e) => {
                          setHAlign(e.target.value as CellStyle['hAlign'])
                          modified.current.add('hAlign')
                        }}
                      >
                        <option value="general">General</option>
                        <option value="left">Left (Indent)</option>
                        <option value="center">Center</option>
                        <option value="right">Right (Indent)</option>
                        <option value="fill">Fill</option>
                        <option value="justify">Justify</option>
                        <option value="centerContinuous">Center Across Selection</option>
                        <option value="distributed">Distributed (Indent)</option>
                      </select>
                    </div>
                    <div className="xl-fc-row" style={{ justifyContent: 'space-between' }}>
                      <span>Indent:</span>
                      <input
                        type="number"
                        min="0"
                        max="250"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={indent}
                        disabled={hAlign !== 'left' && hAlign !== 'right' && hAlign !== 'distributed'}
                        onChange={(e) => {
                          setIndent(Math.max(0, +e.target.value))
                          modified.current.add('indent')
                        }}
                      />
                    </div>
                    <div className="xl-fc-row" style={{ justifyContent: 'space-between' }}>
                      <span>Vertical:</span>
                      <select
                        className="xl-fc-select"
                        style={{ width: 170 }}
                        value={vAlign}
                        onChange={(e) => {
                          setVAlign(e.target.value as CellStyle['vAlign'])
                          modified.current.add('vAlign')
                        }}
                      >
                        <option value="top">Top</option>
                        <option value="middle">Center</option>
                        <option value="bottom">Bottom</option>
                        <option value="justify">Justify</option>
                        <option value="distributed">Distributed</option>
                      </select>
                    </div>
                    <label className="xl-fc-checkbox-label" style={{ marginTop: 2 }}>
                      <input
                        type="checkbox"
                        checked={justifyDistributed}
                        disabled={hAlign !== 'distributed' && hAlign !== 'justify'}
                        onChange={(e) => setJustifyDistributed(e.target.checked)}
                      />
                      <span>Justify distributed</span>
                    </label>
                  </div>
                </div>

                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Text control</span>
                  <div className="xl-fc-col" style={{ gap: 8 }}>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={wrap}
                        onChange={(e) => {
                          setWrap(e.target.checked)
                          modified.current.add('wrap')
                        }}
                      />
                      <span>Wrap text</span>
                    </label>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={shrinkToFit}
                        onChange={(e) => {
                          setShrinkToFit(e.target.checked)
                          modified.current.add('shrinkToFit')
                        }}
                      />
                      <span>Shrink to fit</span>
                    </label>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={merged}
                        onChange={(e) => {
                          setMerged(e.target.checked)
                          modified.current.add('merged')
                        }}
                      />
                      <span>Merge cells</span>
                    </label>
                  </div>
                </div>

                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Right-to-left</span>
                  <div className="xl-fc-row" style={{ justifyContent: 'space-between' }}>
                    <span>Text direction:</span>
                    <select className="xl-fc-select" style={{ width: 140 }}>
                      <option>Context</option>
                      <option>Left-to-Right</option>
                      <option>Right-to-Left</option>
                    </select>
                  </div>
                </div>
              </div>

              <div className="xl-fc-align-right">
                <div className="xl-fc-group" style={{ height: '94%' }}>
                  <span className="xl-fc-group-title">Orientation</span>
                  <div className="xl-fc-col" style={{ alignItems: 'center', gap: 12, paddingTop: 10 }}>
                    <div className="xl-fc-orientation-wrap">
                      {/* Vertical Text Button */}
                      <div
                        className={`xl-fc-vert-text-btn ${rotation === 255 ? 'active' : ''}`}
                        onClick={() => {
                          const next = rotation === 255 ? 0 : 255
                          setRotation(next)
                          modified.current.add('rotation')
                        }}
                        title="Toggle Vertical Text"
                      >
                        <span>T</span>
                        <span>e</span>
                        <span>x</span>
                        <span>t</span>
                      </div>

                      {/* Half-circle Dial Canvas */}
                      <div className="xl-fc-dial-canvas-wrap">
                        <canvas
                          ref={dialRef}
                          width={110}
                          height={110}
                          onMouseDown={handleDialMouseDown}
                        />
                      </div>
                    </div>

                    <div className="xl-fc-row">
                      <span>Degrees:</span>
                      <input
                        type="number"
                        min="-90"
                        max="90"
                        className="xl-fc-input"
                        style={{ width: 60 }}
                        value={rotation === 255 ? 0 : rotation}
                        onChange={(e) => {
                          const deg = Math.max(-90, Math.min(90, +e.target.value))
                          setRotation(deg)
                          modified.current.add('rotation')
                        }}
                      />
                      <span>°</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: FONT */}
          {tab === 'Font' && (
            <div className="xl-fc-font-layout">
              <div className="xl-fc-font-cols">
                {/* Font Family */}
                <div className="xl-fc-font-col" style={{ flex: 1.6 }}>
                  <span>Font:</span>
                  <input
                    className="xl-fc-input"
                    value={fontFamilyInput}
                    onChange={(e) => {
                      setFontFamilyInput(e.target.value)
                      const found = FONT_FAMILIES.find((f) => f.toLowerCase() === e.target.value.toLowerCase())
                      if (found) {
                        setFontFamily(found)
                        modified.current.add('fontFamily')
                      }
                    }}
                  />
                  <div className="xl-fc-listbox" style={{ height: 110 }}>
                    {FONT_FAMILIES.filter((fn) =>
                      fn.toLowerCase().includes(fontFamilyInput.toLowerCase())
                    ).map((fn) => (
                      <div
                        key={fn}
                        className={`xl-fc-list-item ${fontFamily.toLowerCase() === fn.toLowerCase() ? 'selected' : ''}`}
                        onClick={() => {
                          setFontFamily(fn)
                          setFontFamilyInput(fn)
                          modified.current.add('fontFamily')
                        }}
                      >
                        {fn}
                      </div>
                    ))}
                  </div>
                </div>

                {/* Font Style */}
                <div className="xl-fc-font-col" style={{ flex: 1.2 }}>
                  <span>Font style:</span>
                  <input
                    className="xl-fc-input"
                    value={fontStyleInput}
                    readOnly
                  />
                  <div className="xl-fc-listbox" style={{ height: 110 }}>
                    {[
                      { l: 'Regular', b: false, i: false },
                      { l: 'Italic', b: false, i: true },
                      { l: 'Bold', b: true, i: false },
                      { l: 'Bold Italic', b: true, i: true }
                    ].map((s) => (
                      <div
                        key={s.l}
                        className={`xl-fc-list-item ${bold === s.b && italic === s.i ? 'selected' : ''}`}
                        onClick={() => {
                          setBold(s.b)
                          setItalic(s.i)
                          setFontStyleInput(s.l)
                          modified.current.add('bold')
                          modified.current.add('italic')
                        }}
                      >
                        {s.l}
                      </div>
                    ))}
                  </div>
                </div>

                {/* Size */}
                <div className="xl-fc-font-col" style={{ flex: 0.8 }}>
                  <span>Size:</span>
                  <input
                    className="xl-fc-input"
                    type="number"
                    value={fontSizeInput}
                    onChange={(e) => {
                      setFontSizeInput(e.target.value)
                      const sz = +e.target.value
                      if (sz > 0) {
                        setFontSize(sz)
                        modified.current.add('fontSize')
                      }
                    }}
                  />
                  <div className="xl-fc-listbox" style={{ height: 110 }}>
                    {FONT_SIZES.map((sz) => (
                      <div
                        key={sz}
                        className={`xl-fc-list-item ${fontSize === sz ? 'selected' : ''}`}
                        onClick={() => {
                          setFontSize(sz)
                          setFontSizeInput(String(sz))
                          modified.current.add('fontSize')
                        }}
                      >
                        {sz}
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Underline, Color, Normal Font */}
              <div className="xl-fc-row" style={{ gap: 20, marginTop: 4 }}>
                <div className="xl-fc-row">
                  <span>Underline:</span>
                  <select
                    className="xl-fc-select"
                    style={{ width: 140 }}
                    value={underline}
                    onChange={(e) => {
                      setUnderline(e.target.value)
                      modified.current.add('underline')
                    }}
                  >
                    <option value="none">None</option>
                    <option value="single">Single</option>
                    <option value="double">Double</option>
                    <option value="singleAccounting">Single Accounting</option>
                    <option value="doubleAccounting">Double Accounting</option>
                  </select>
                </div>

                <div className="xl-fc-row" style={{ position: 'relative' }}>
                  <span>Color:</span>
                  <button
                    className="xl-btn"
                    style={{ minWidth: 60, display: 'flex', alignItems: 'center', gap: 4 }}
                    onClick={() => setFontColorOpen(!fontColorOpen)}
                  >
                    <span
                      style={{
                        width: 14,
                        height: 14,
                        background: fontColor || '#000000',
                        border: '1px solid #777'
                      }}
                    />
                    <span>▼</span>
                  </button>

                  {fontColorOpen && (
                    <div
                      style={{
                        position: 'absolute',
                        left: 45,
                        top: 26,
                        background: '#fff',
                        border: '1px solid #adadad',
                        boxShadow: '0 4px 12px rgba(0,0,0,0.2)',
                        padding: 8,
                        zIndex: 100
                      }}
                    >
                      <button
                        className="xl-btn"
                        style={{ width: '100%', marginBottom: 6 }}
                        onClick={() => {
                          setFontColor(undefined)
                          modified.current.add('color')
                          setFontColorOpen(false)
                        }}
                      >
                        Automatic
                      </button>
                      <div className="xl-fc-swatch-grid">
                        {themeRows().map((row, ri) =>
                          row.map((c, ci) => (
                            <div
                              key={`${ri}-${ci}`}
                              className="xl-fc-swatch"
                              style={{ background: c }}
                              onClick={() => {
                                setFontColor(c)
                                modified.current.add('color')
                                setFontColorOpen(false)
                              }}
                            />
                          ))
                        )}
                      </div>
                    </div>
                  )}
                </div>

                <label className="xl-fc-checkbox-label">
                  <input
                    type="checkbox"
                    checked={fontFamily === 'Aptos Narrow' && fontSize === 11 && !bold && !italic && underline === 'none' && !fontColor}
                    onChange={(e) => {
                      if (e.target.checked) {
                        setFontFamily('Aptos Narrow')
                        setFontFamilyInput('Aptos Narrow')
                        setFontSize(11)
                        setFontSizeInput('11')
                        setBold(false)
                        setItalic(false)
                        setFontStyleInput('Regular')
                        setUnderline('none')
                        setFontColor(undefined)
                        setStrike(false)
                        setSuperscript(false)
                        setSubscript(false)
                        modified.current.add('fontFamily')
                        modified.current.add('fontSize')
                        modified.current.add('bold')
                        modified.current.add('italic')
                        modified.current.add('underline')
                        modified.current.add('color')
                        modified.current.add('strike')
                      }
                    }}
                  />
                  <span>Normal font</span>
                </label>
              </div>

              {/* Effects & Preview */}
              <div className="xl-fc-row" style={{ gap: 16, marginTop: 4, alignItems: 'flex-start' }}>
                <div className="xl-fc-group" style={{ flex: 1 }}>
                  <span className="xl-fc-group-title">Effects</span>
                  <div className="xl-fc-col" style={{ gap: 6 }}>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={strike}
                        onChange={(e) => {
                          setStrike(e.target.checked)
                          modified.current.add('strike')
                        }}
                      />
                      <span>Strikethrough</span>
                    </label>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={superscript}
                        onChange={(e) => {
                          setSuperscript(e.target.checked)
                          if (e.target.checked) setSubscript(false)
                        }}
                      />
                      <span>Superscript</span>
                    </label>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={subscript}
                        onChange={(e) => {
                          setSubscript(e.target.checked)
                          if (e.target.checked) setSuperscript(false)
                        }}
                      />
                      <span>Subscript</span>
                    </label>
                  </div>
                </div>

                <div className="xl-fc-group" style={{ flex: 1.5 }}>
                  <span className="xl-fc-group-title">Preview</span>
                  <div className="xl-fc-font-preview">
                    <span
                      style={{
                        fontFamily,
                        fontSize: Math.min(fontSize, 22),
                        fontWeight: bold ? 'bold' : 'normal',
                        fontStyle: italic ? 'italic' : 'normal',
                        textDecoration: [
                          underline !== 'none' ? 'underline' : '',
                          strike ? 'line-through' : ''
                        ]
                          .filter(Boolean)
                          .join(' '),
                        color: fontColor || '#000'
                      }}
                    >
                      AaBbCcYyZz
                    </span>
                  </div>
                </div>
              </div>

              <div className="xl-fc-help-text">
                This is a TrueType font. The same font will be used on both your printer and your screen.
              </div>
            </div>
          )}

          {/* TAB 4: BORDER */}
          {tab === 'Border' && (
            <div className="xl-fc-border-layout">
              {/* Left Column: Style & Color */}
              <div className="xl-fc-border-left">
                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Line Style</span>
                  <div className="xl-fc-style-grid">
                    {LINE_STYLES.map((ls) => (
                      <div
                        key={ls.id}
                        className={`xl-fc-style-cell ${selectedStyleId === ls.id ? 'selected' : ''}`}
                        onClick={() => setSelectedStyleId(ls.id)}
                        title={ls.label}
                      >
                        {ls.id === 0 ? (
                          <span style={{ fontSize: 11, color: '#666' }}>None</span>
                        ) : (
                          <div style={{ width: '100%', height: 0, borderTop: ls.borderCss }} />
                        )}
                      </div>
                    ))}
                  </div>
                </div>

                <div className="xl-fc-row">
                  <span>Color:</span>
                  <input
                    type="color"
                    className="xl-fc-input"
                    style={{ width: 44, padding: 0, cursor: 'pointer' }}
                    value={borderColor}
                    onChange={(e) => setBorderColor(e.target.value)}
                  />
                  <button
                    className="xl-btn"
                    style={{ minWidth: 60 }}
                    onClick={() => setBorderColor('#000000')}
                  >
                    Automatic
                  </button>
                </div>
              </div>

              {/* Right Column: Presets & 8 Toggle Buttons + Diagram */}
              <div className="xl-fc-border-right">
                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Presets</span>
                  <div className="xl-fc-presets-row">
                    <button className="xl-fc-preset-btn" onClick={() => handleBorderPreset('none')}>
                      <span>✕</span>
                      <span>None</span>
                    </button>
                    <button className="xl-fc-preset-btn" onClick={() => handleBorderPreset('outline')}>
                      <span>▢</span>
                      <span>Outline</span>
                    </button>
                    <button
                      className="xl-fc-preset-btn"
                      disabled={!isMultiCell}
                      onClick={() => handleBorderPreset('inside')}
                      title={isMultiCell ? 'Inside' : 'Inside (available when several cells are selected)'}
                    >
                      <span>⊞</span>
                      <span>Inside</span>
                    </button>
                  </div>
                </div>

                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Border</span>
                  <div className="xl-fc-diagram-container">
                    <div className="xl-fc-diagram-middle">
                      {/* Left column of 4 toggle buttons */}
                      <div className="xl-fc-side-btns">
                        <button
                          className={`xl-fc-border-toolbtn ${borders.top ? 'active' : ''}`}
                          onClick={() => toggleBorder('top')}
                          title="Top Border"
                        >
                          ▔
                        </button>
                        <button
                          className={`xl-fc-border-toolbtn ${borders.insideH ? 'active' : ''}`}
                          disabled={!isMultiCell}
                          onClick={() => toggleBorder('insideH')}
                          title="Inside Horizontal Border"
                        >
                          ―
                        </button>
                        <button
                          className={`xl-fc-border-toolbtn ${borders.bottom ? 'active' : ''}`}
                          onClick={() => toggleBorder('bottom')}
                          title="Bottom Border"
                        >
                           
                        </button>
                        <button
                          className={`xl-fc-border-toolbtn ${borders.diagUp ? 'active' : ''}`}
                          onClick={() => toggleBorder('diagUp')}
                          title="Diagonal Up"
                        >
                          ╱
                        </button>
                      </div>

                      {/* Interactive Diagram Box */}
                      <div
                        className={`xl-fc-border-diagram-box ${isMultiCell ? 'quad' : ''}`}
                        style={{
                          borderTop: borders.top ? `3px solid ${borders.top.color || '#000'}` : '1px dashed #adadad',
                          borderBottom: borders.bottom ? `3px solid ${borders.bottom.color || '#000'}` : '1px dashed #adadad',
                          borderLeft: borders.left ? `3px solid ${borders.left.color || '#000'}` : '1px dashed #adadad',
                          borderRight: borders.right ? `3px solid ${borders.right.color || '#000'}` : '1px dashed #adadad'
                        }}
                      >
                        {isMultiCell ? (
                          <>
                            <div
                              className="xl-fc-border-quadrant"
                              style={{
                                borderBottom: borders.insideH ? `2px solid ${borders.insideH.color || '#000'}` : '1px dashed #ddd',
                                borderRight: borders.insideV ? `2px solid ${borders.insideV.color || '#000'}` : '1px dashed #ddd'
                              }}
                            >
                              Text
                            </div>
                            <div
                              className="xl-fc-border-quadrant"
                              style={{
                                borderBottom: borders.insideH ? `2px solid ${borders.insideH.color || '#000'}` : '1px dashed #ddd'
                              }}
                            >
                              Text
                            </div>
                            <div
                              className="xl-fc-border-quadrant"
                              style={{
                                borderRight: borders.insideV ? `2px solid ${borders.insideV.color || '#000'}` : '1px dashed #ddd'
                              }}
                            >
                              Text
                            </div>
                            <div className="xl-fc-border-quadrant">Text</div>
                          </>
                        ) : (
                          <div className="xl-fc-border-quadrant" style={{ width: '100%', height: '100%' }}>
                            Text
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Bottom row of 4 toggle buttons */}
                    <div className="xl-fc-bottom-btns">
                      <button
                        className={`xl-fc-border-toolbtn ${borders.left ? 'active' : ''}`}
                        onClick={() => toggleBorder('left')}
                        title="Left Border"
                      >
                        ▏
                      </button>
                      <button
                        className={`xl-fc-border-toolbtn ${borders.insideV ? 'active' : ''}`}
                        disabled={!isMultiCell}
                        onClick={() => toggleBorder('insideV')}
                        title="Inside Vertical Border"
                      >
                        │
                      </button>
                      <button
                        className={`xl-fc-border-toolbtn ${borders.right ? 'active' : ''}`}
                        onClick={() => toggleBorder('right')}
                        title="Right Border"
                      >
                        ▕
                      </button>
                      <button
                        className={`xl-fc-border-toolbtn ${borders.diagDown ? 'active' : ''}`}
                        onClick={() => toggleBorder('diagDown')}
                        title="Diagonal Down"
                      >
                        ╲
                      </button>
                    </div>
                  </div>
                </div>

                <div className="xl-fc-help-text">
                  The selected border style can be applied by clicking the presets, preview diagram or the buttons above.
                </div>
              </div>
            </div>
          )}

          {/* TAB 5: FILL */}
          {tab === 'Fill' && (
            <div className="xl-fc-fill-layout">
              <div className="xl-fc-row" style={{ justifyContent: 'space-between' }}>
                <span style={{ fontWeight: 500 }}>Background Color</span>
                <button
                  className="xl-btn"
                  onClick={() => {
                    setFill(undefined)
                    modified.current.add('fill')
                  }}
                >
                  No Color
                </button>
              </div>

              <div className="xl-fc-color-palette">
                <div style={{ fontSize: 11, color: '#555' }}>Theme Colors</div>
                <div className="xl-fc-swatch-grid">
                  {themeRows().map((row, ri) =>
                    row.map((c, ci) => (
                      <div
                        key={`${ri}-${ci}`}
                        className={`xl-fc-swatch ${fill?.toUpperCase() === c.toUpperCase() ? 'selected' : ''}`}
                        style={{ background: c }}
                        title={c}
                        onClick={() => {
                          setFill(c)
                          modified.current.add('fill')
                        }}
                      />
                    ))
                  )}
                </div>

                <div style={{ fontSize: 11, color: '#555', marginTop: 4 }}>Standard Colors</div>
                <div className="xl-fc-swatch-grid">
                  {STANDARD.map((c) => (
                    <div
                      key={c}
                      className={`xl-fc-swatch ${fill?.toUpperCase() === c.toUpperCase() ? 'selected' : ''}`}
                      style={{ background: c }}
                      title={c}
                      onClick={() => {
                        setFill(c)
                        modified.current.add('fill')
                      }}
                    />
                  ))}
                </div>
              </div>

              <div className="xl-fc-row" style={{ gap: 10, marginTop: 4 }}>
                <button className="xl-btn" style={{ minWidth: 90 }} disabled>
                  Fill Effects...
                </button>
                <div className="xl-fc-row">
                  <span style={{ fontSize: 11 }}>More Colors:</span>
                  <input
                    type="color"
                    className="xl-fc-input"
                    style={{ width: 44, padding: 0, cursor: 'pointer' }}
                    value={fill || '#FFFFFF'}
                    onChange={(e) => {
                      setFill(e.target.value)
                      modified.current.add('fill')
                    }}
                  />
                </div>
              </div>

              <div className="xl-fc-row" style={{ gap: 20, marginTop: 4 }}>
                <div className="xl-fc-row">
                  <span>Pattern Color:</span>
                  <select
                    className="xl-fc-select"
                    value={patternColor}
                    onChange={(e) => setPatternColor(e.target.value)}
                  >
                    <option value="#000000">Automatic</option>
                  </select>
                </div>

                <div className="xl-fc-row">
                  <span>Pattern Style:</span>
                  <select
                    className="xl-fc-select"
                    value={patternStyle}
                    onChange={(e) => setPatternStyle(e.target.value)}
                  >
                    <option value="Automatic">Automatic</option>
                    <option value="5% Gray">5% Gray</option>
                    <option value="10% Gray">10% Gray</option>
                    <option value="25% Gray">25% Gray</option>
                    <option value="50% Gray">50% Gray</option>
                    <option value="Thin Horizontal">Thin Horizontal</option>
                    <option value="Thin Vertical">Thin Vertical</option>
                  </select>
                </div>
              </div>

              <div className="xl-fc-group" style={{ marginTop: 4 }}>
                <span className="xl-fc-group-title">Sample</span>
                <div
                  style={{
                    height: 38,
                    border: '1px solid #adadad',
                    background: fill || '#ffffff',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    color: fill ? '#000000' : '#888'
                  }}
                >
                  {fill ? 'Sample' : 'No Fill'}
                </div>
              </div>
            </div>
          )}

          {/* TAB 6: PROTECTION */}
          {tab === 'Protection' && (
            <div className="xl-fc-prot-layout">
              <label className="xl-fc-checkbox-label">
                <input
                  type="checkbox"
                  checked={locked}
                  onChange={(e) => {
                    setLocked(e.target.checked)
                    modified.current.add('locked')
                  }}
                />
                <span>Locked</span>
              </label>

              <label className="xl-fc-checkbox-label">
                <input
                  type="checkbox"
                  checked={hidden}
                  onChange={(e) => {
                    setHidden(e.target.checked)
                    modified.current.add('hidden')
                  }}
                />
                <span>Hidden</span>
              </label>

              <div className="xl-fc-help-text" style={{ marginTop: 12 }}>
                Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect Sheet button).
              </div>
            </div>
          )}
        </div>
      </div>
    </Dialog>
  )
}

import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import {
  BorderLineStyle,
  BorderSide,
  CellStyle,
  CellValue
} from '../model/types'
import {
  FORMAT_CATEGORIES,
  buildNumberFormat,
  formatValue
} from '../format/numfmt'
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

const THEME_BASE = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
function shade(hex: string, amt: number): string {
  const n = parseInt(hex.slice(1), 16)
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255
  const mix = (c: number) => (amt > 0 ? c + (255 - c) * amt : c * (1 + amt))
  r = Math.round(mix(r)); g = Math.round(mix(g)); b = Math.round(mix(b))
  return '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase()
}
const THEME_ROWS = [0, 0.8, 0.6, 0.4, -0.25, -0.5].map((amt, i) => (i === 0 ? THEME_BASE : THEME_BASE.map((c) => shade(c, amt))))
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0']

const FONT_FAMILIES = [
  'Aptos',
  'Aptos Narrow',
  'Arial',
  'Calibri',
  'Cambria',
  'Comic Sans MS',
  'Consolas',
  'Courier New',
  'Georgia',
  'Impact',
  'Segoe UI',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana'
]

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

const BORDER_STYLES: { label: string; style: BorderLineStyle; previewBorder: string }[] = [
  { label: 'None', style: 'thin', previewBorder: 'none' },
  { label: 'Thin', style: 'thin', previewBorder: '1px solid #000' },
  { label: 'Medium', style: 'medium', previewBorder: '2px solid #000' },
  { label: 'Dashed', style: 'dashed', previewBorder: '1px dashed #000' },
  { label: 'Dotted', style: 'dotted', previewBorder: '1px dotted #000' },
  { label: 'Thick', style: 'thick', previewBorder: '3px solid #000' },
  { label: 'Double', style: 'double', previewBorder: '3px double #000' },
  { label: 'Hair', style: 'hair', previewBorder: '1px dotted #888' },
  { label: 'Medium Dashed', style: 'mediumDashed', previewBorder: '2px dashed #000' },
  { label: 'Dash Dot', style: 'dashDot', previewBorder: '1px dashed #000' },
  { label: 'Medium Dash Dot', style: 'mediumDashDot', previewBorder: '2px dashed #000' },
  { label: 'Dash Dot Dot', style: 'dashDotDot', previewBorder: '1px dotted #000' },
  { label: 'Slant Dash Dot', style: 'slantDashDot', previewBorder: '2px dashed #000' }
]

export function FormatCellsDialog({
  initialTab = 'Number',
  initialStyle = {},
  cellValue,
  onApply,
  onClose
}: FormatCellsDialogProps) {
  const [tab, setTab] = useState<FormatCellsTab>(
    (TABS.find((t) => t.toLowerCase() === initialTab.toLowerCase()) as FormatCellsTab) || 'Number'
  )

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
  const [currencySymbol, setCurrencySymbol] = useState<string>('$')

  // --- Alignment State ---
  const [hAlign, setHAlign] = useState<CellStyle['hAlign']>(initialStyle.hAlign || 'general')
  const [vAlign, setVAlign] = useState<CellStyle['vAlign']>(initialStyle.vAlign || 'bottom')
  const [indent, setIndent] = useState<number>(initialStyle.indent || 0)
  const [wrap, setWrap] = useState<boolean>(!!initialStyle.wrap)
  const [shrinkToFit, setShrinkToFit] = useState<boolean>(!!initialStyle.shrinkToFit)
  const [rotation, setRotation] = useState<number>(initialStyle.rotation ?? 0)

  // --- Font State ---
  const [fontFamily, setFontFamily] = useState<string>(initialStyle.fontFamily || 'Aptos Narrow')
  const [fontSize, setFontSize] = useState<number>(initialStyle.fontSize || 11)
  const [bold, setBold] = useState<boolean>(!!initialStyle.bold)
  const [italic, setItalic] = useState<boolean>(!!initialStyle.italic)
  const [underline, setUnderline] = useState<'single' | 'double' | 'none'>(initialStyle.underline || 'none')
  const [strike, setStrike] = useState<boolean>(!!initialStyle.strike)
  const [fontColor, setFontColor] = useState<string | undefined>(initialStyle.color)

  // --- Border State ---
  const [borders, setBorders] = useState<{
    top?: BorderSide
    right?: BorderSide
    bottom?: BorderSide
    left?: BorderSide
  }>(initialStyle.borders || {})
  const [activeBorderStyle, setActiveBorderStyle] = useState<BorderLineStyle>('thin')
  const [activeBorderColor, setActiveBorderColor] = useState<string>('#000000')

  // --- Fill State ---
  const [fill, setFill] = useState<string | undefined>(initialStyle.fill)

  // --- Protection State ---
  const [locked, setLocked] = useState<boolean>(initialStyle.locked !== false)
  const [hidden, setHidden] = useState<boolean>(!!initialStyle.hidden)

  // Handlers for Number tab
  const handleCategorySelect = (cat: string) => {
    setNumCategory(cat)
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

  const handleNumberOptionsChange = (decimals: number, thousands: boolean, negative: 0 | 1 | 2 | 3) => {
    setNumDecimals(decimals)
    setNumThousands(thousands)
    setNumNegative(negative)
    setNumCode(buildNumberFormat({ decimals, thousands, negative }))
  }

  const handleCurrencyOptionsChange = (decimals: number, symbol: string, negative: 0 | 1 | 2 | 3) => {
    setNumDecimals(decimals)
    setCurrencySymbol(symbol)
    setNumNegative(negative)
    setNumCode(buildNumberFormat({ decimals, thousands: true, negative, symbol: symbol === 'None' ? undefined : symbol }))
  }

  // Format sample value
  const sampleVal = cellValue !== undefined && cellValue !== null && cellValue !== '' ? cellValue : 1234.1
  const sampleDisplay = formatValue(sampleVal, numCode === 'General' ? undefined : numCode).text

  // Border toggle
  const toggleBorder = (side: 'top' | 'right' | 'bottom' | 'left') => {
    setBorders((prev) => {
      const current = prev[side]
      if (current) {
        const next = { ...prev }
        delete next[side]
        return next
      }
      return {
        ...prev,
        [side]: { style: activeBorderStyle, color: activeBorderColor }
      }
    })
  }

  const setPresetBorder = (preset: 'none' | 'outline') => {
    if (preset === 'none') {
      setBorders({})
    } else if (preset === 'outline') {
      setBorders({
        top: { style: activeBorderStyle, color: activeBorderColor },
        bottom: { style: activeBorderStyle, color: activeBorderColor },
        left: { style: activeBorderStyle, color: activeBorderColor },
        right: { style: activeBorderStyle, color: activeBorderColor }
      })
    }
  }

  const handleOk = () => {
    const patch: Partial<CellStyle> = {
      numFmt: numCode,
      hAlign: hAlign === 'general' ? undefined : hAlign,
      vAlign: vAlign === 'bottom' ? undefined : vAlign,
      indent: indent > 0 ? indent : undefined,
      wrap: wrap || undefined,
      shrinkToFit: shrinkToFit || undefined,
      rotation: rotation !== 0 ? rotation : undefined,
      fontFamily: fontFamily !== 'Aptos Narrow' ? fontFamily : undefined,
      fontSize: fontSize !== 11 ? fontSize : undefined,
      bold: bold || undefined,
      italic: italic || undefined,
      underline: underline === 'none' ? undefined : underline,
      strike: strike || undefined,
      color: fontColor,
      borders: Object.keys(borders).length > 0 ? borders : undefined,
      fill: fill || undefined,
      locked: locked ? undefined : false,
      hidden: hidden ? true : undefined
    }
    onApply(patch)
    onClose()
  }

  return (
    <Dialog title="Format Cells" width={600} onOk={handleOk} onCancel={onClose}>
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
                <div className="xl-fc-listbox" style={{ height: 320 }}>
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
                        onChange={(e) => handleNumberOptionsChange(Math.max(0, +e.target.value), numThousands, numNegative)}
                      />
                    </div>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={numThousands}
                        onChange={(e) => handleNumberOptionsChange(numDecimals, e.target.checked, numNegative)}
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
                            onClick={() => handleNumberOptionsChange(numDecimals, numThousands, opt.val as 0 | 1 | 2 | 3)}
                          >
                            {opt.text}
                          </div>
                        ))}
                      </div>
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
                        onChange={(e) => handleCurrencyOptionsChange(Math.max(0, +e.target.value), currencySymbol, numNegative)}
                      />
                    </div>
                    <div className="xl-fc-row">
                      <span>Symbol:</span>
                      <select
                        className="xl-fc-select"
                        value={currencySymbol}
                        onChange={(e) => handleCurrencyOptionsChange(numDecimals, e.target.value, numNegative)}
                      >
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
                            onClick={() => handleCurrencyOptionsChange(numDecimals, currencySymbol, opt.val as 0 | 1 | 2 | 3)}
                          >
                            {opt.text}
                          </div>
                        ))}
                      </div>
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
                        onChange={(e) => setNumDecimals(Math.max(0, +e.target.value))}
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
                        }}
                      >
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
                          onClick={() => setNumCode(entry.code)}
                        >
                          {entry.label}
                        </div>
                      ))}
                    </div>
                    <div className="xl-fc-row" style={{ marginTop: 8 }}>
                      <span>Locale (location):</span>
                      <select className="xl-fc-select" style={{ width: 180 }} disabled>
                        <option>English (United States)</option>
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
                        }}
                      />
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
                        onChange={(e) => setNumCode(e.target.value)}
                      />
                    </div>
                    <div className="xl-fc-listbox" style={{ height: 180, marginTop: 4 }}>
                      {FORMAT_CATEGORIES.Custom.map((entry) => (
                        <div
                          key={entry.code}
                          className={`xl-fc-list-item ${numCode === entry.code ? 'selected' : ''}`}
                          onClick={() => setNumCode(entry.code)}
                        >
                          {entry.code}
                        </div>
                      ))}
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
                        style={{ width: 160 }}
                        value={hAlign}
                        onChange={(e) => setHAlign(e.target.value as CellStyle['hAlign'])}
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
                        onChange={(e) => setIndent(Math.max(0, +e.target.value))}
                      />
                    </div>
                    <div className="xl-fc-row" style={{ justifyContent: 'space-between' }}>
                      <span>Vertical:</span>
                      <select
                        className="xl-fc-select"
                        style={{ width: 160 }}
                        value={vAlign}
                        onChange={(e) => setVAlign(e.target.value as CellStyle['vAlign'])}
                      >
                        <option value="top">Top</option>
                        <option value="middle">Center</option>
                        <option value="bottom">Bottom</option>
                        <option value="justify">Justify</option>
                        <option value="distributed">Distributed</option>
                      </select>
                    </div>
                  </div>
                </div>

                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Text control</span>
                  <div className="xl-fc-col" style={{ gap: 8 }}>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={wrap}
                        onChange={(e) => setWrap(e.target.checked)}
                      />
                      <span>Wrap text</span>
                    </label>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={shrinkToFit}
                        onChange={(e) => setShrinkToFit(e.target.checked)}
                      />
                      <span>Shrink to fit</span>
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
                <div className="xl-fc-group" style={{ height: '90%' }}>
                  <span className="xl-fc-group-title">Orientation</span>
                  <div className="xl-fc-col" style={{ alignItems: 'center', gap: 14, paddingTop: 10 }}>
                    <div className="xl-fc-orientation-box">
                      <div
                        className={`xl-fc-vert-btn ${rotation === 255 ? 'active' : ''}`}
                        onClick={() => setRotation(rotation === 255 ? 0 : 255)}
                        title="Vertical Text"
                      >
                        <span>T</span>
                        <span>e</span>
                        <span>x</span>
                        <span>t</span>
                      </div>
                      <div
                        style={{
                          width: 80,
                          height: 80,
                          border: '1px solid #d4d4d4',
                          borderRadius: '50%',
                          position: 'relative',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          background: '#fafafa'
                        }}
                      >
                        <div
                          style={{
                            width: 30,
                            height: 2,
                            background: '#0078d7',
                            transformOrigin: '0 50%',
                            transform: `rotate(${-rotation}deg)`
                          }}
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
                        onChange={(e) => setRotation(Math.max(-90, Math.min(90, +e.target.value)))}
                      />
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
                <div className="xl-fc-font-col" style={{ flex: 1.6 }}>
                  <span>Font:</span>
                  <input
                    className="xl-fc-input"
                    value={fontFamily}
                    onChange={(e) => setFontFamily(e.target.value)}
                  />
                  <div className="xl-fc-listbox" style={{ height: 110 }}>
                    {FONT_FAMILIES.map((fn) => (
                      <div
                        key={fn}
                        className={`xl-fc-list-item ${fontFamily.toLowerCase() === fn.toLowerCase() ? 'selected' : ''}`}
                        onClick={() => setFontFamily(fn)}
                      >
                        {fn}
                      </div>
                    ))}
                  </div>
                </div>

                <div className="xl-fc-font-col" style={{ flex: 1.2 }}>
                  <span>Font style:</span>
                  <input
                    className="xl-fc-input"
                    value={
                      bold && italic
                        ? 'Bold Italic'
                        : bold
                          ? 'Bold'
                          : italic
                            ? 'Italic'
                            : 'Regular'
                    }
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
                        }}
                      >
                        {s.l}
                      </div>
                    ))}
                  </div>
                </div>

                <div className="xl-fc-font-col" style={{ flex: 0.8 }}>
                  <span>Size:</span>
                  <input
                    className="xl-fc-input"
                    type="number"
                    value={fontSize}
                    onChange={(e) => setFontSize(Math.max(1, +e.target.value))}
                  />
                  <div className="xl-fc-listbox" style={{ height: 110 }}>
                    {FONT_SIZES.map((sz) => (
                      <div
                        key={sz}
                        className={`xl-fc-list-item ${fontSize === sz ? 'selected' : ''}`}
                        onClick={() => setFontSize(sz)}
                      >
                        {sz}
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="xl-fc-row" style={{ gap: 24, marginTop: 4 }}>
                <div className="xl-fc-row">
                  <span>Underline:</span>
                  <select
                    className="xl-fc-select"
                    style={{ width: 110 }}
                    value={underline}
                    onChange={(e) => setUnderline(e.target.value as 'single' | 'double' | 'none')}
                  >
                    <option value="none">None</option>
                    <option value="single">Single</option>
                    <option value="double">Double</option>
                  </select>
                </div>

                <div className="xl-fc-row">
                  <span>Color:</span>
                  <input
                    type="color"
                    className="xl-fc-input"
                    style={{ width: 44, padding: 0, cursor: 'pointer' }}
                    value={fontColor || '#000000'}
                    onChange={(e) => setFontColor(e.target.value)}
                  />
                  {fontColor && (
                    <button
                      className="xl-btn"
                      style={{ minWidth: 60 }}
                      onClick={() => setFontColor(undefined)}
                    >
                      Automatic
                    </button>
                  )}
                </div>
              </div>

              <div className="xl-fc-row" style={{ gap: 16, marginTop: 4, alignItems: 'flex-start' }}>
                <div className="xl-fc-group" style={{ flex: 1 }}>
                  <span className="xl-fc-group-title">Effects</span>
                  <div className="xl-fc-col" style={{ gap: 8 }}>
                    <label className="xl-fc-checkbox-label">
                      <input
                        type="checkbox"
                        checked={strike}
                        onChange={(e) => setStrike(e.target.checked)}
                      />
                      <span>Strikethrough</span>
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
              <div className="xl-fc-border-left">
                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Line Style</span>
                  <div className="xl-fc-listbox" style={{ height: 180 }}>
                    {BORDER_STYLES.map((b) => (
                      <div
                        key={b.label}
                        className={`xl-fc-list-item ${activeBorderStyle === b.style ? 'selected' : ''}`}
                        onClick={() => setActiveBorderStyle(b.style)}
                        style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}
                      >
                        <span>{b.label}</span>
                        <div style={{ width: 40, height: 4, borderTop: b.previewBorder }} />
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
                    value={activeBorderColor}
                    onChange={(e) => setActiveBorderColor(e.target.value)}
                  />
                </div>
              </div>

              <div className="xl-fc-border-right">
                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Presets</span>
                  <div className="xl-fc-presets-row">
                    <button className="xl-fc-preset-btn" onClick={() => setPresetBorder('none')}>
                      <span>✕</span>
                      <span>None</span>
                    </button>
                    <button className="xl-fc-preset-btn" onClick={() => setPresetBorder('outline')}>
                      <span>▢</span>
                      <span>Outline</span>
                    </button>
                  </div>
                </div>

                <div className="xl-fc-group">
                  <span className="xl-fc-group-title">Border</span>
                  <div className="xl-fc-col" style={{ alignItems: 'center', gap: 6 }}>
                    {/* Top toggle */}
                    <button
                      className={`xl-fc-border-line-btn ${borders.top ? 'active' : ''}`}
                      onClick={() => toggleBorder('top')}
                      title="Top Border"
                    >
                      ▲
                    </button>
                    <div className="xl-fc-border-diagram-wrap">
                      {/* Left toggle */}
                      <button
                        className={`xl-fc-border-line-btn ${borders.left ? 'active' : ''}`}
                        onClick={() => toggleBorder('left')}
                        title="Left Border"
                      >
                        ◀
                      </button>

                      {/* Cell box preview */}
                      <div
                        className="xl-fc-border-cell-box"
                        style={{
                          borderTop: borders.top
                            ? `3px solid ${borders.top.color || '#000'}`
                            : '1px dashed #bbb',
                          borderBottom: borders.bottom
                            ? `3px solid ${borders.bottom.color || '#000'}`
                            : '1px dashed #bbb',
                          borderLeft: borders.left
                            ? `3px solid ${borders.left.color || '#000'}`
                            : '1px dashed #bbb',
                          borderRight: borders.right
                            ? `3px solid ${borders.right.color || '#000'}`
                            : '1px dashed #bbb'
                        }}
                      >
                        <span style={{ color: '#aaa', fontSize: 13 }}>Text</span>
                      </div>

                      {/* Right toggle */}
                      <button
                        className={`xl-fc-border-line-btn ${borders.right ? 'active' : ''}`}
                        onClick={() => toggleBorder('right')}
                        title="Right Border"
                      >
                        ▶
                      </button>
                    </div>

                    {/* Bottom toggle */}
                    <button
                      className={`xl-fc-border-line-btn ${borders.bottom ? 'active' : ''}`}
                      onClick={() => toggleBorder('bottom')}
                      title="Bottom Border"
                    >
                      ▼
                    </button>
                  </div>
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
                  onClick={() => setFill(undefined)}
                >
                  No Color
                </button>
              </div>

              <div className="xl-fc-color-palette">
                <div style={{ fontSize: 11, color: '#555' }}>Theme Colors</div>
                <div className="xl-fc-swatch-grid">
                  {THEME_ROWS.map((row, ri) =>
                    row.map((c, ci) => (
                      <div
                        key={`${ri}-${ci}`}
                        className={`xl-fc-swatch ${fill?.toUpperCase() === c.toUpperCase() ? 'selected' : ''}`}
                        style={{ background: c }}
                        title={c}
                        onClick={() => setFill(c)}
                      />
                    ))
                  )}
                </div>

                <div style={{ fontSize: 11, color: '#555', marginTop: 6 }}>Standard Colors</div>
                <div className="xl-fc-swatch-grid">
                  {STANDARD.map((c) => (
                    <div
                      key={c}
                      className={`xl-fc-swatch ${fill?.toUpperCase() === c.toUpperCase() ? 'selected' : ''}`}
                      style={{ background: c }}
                      title={c}
                      onClick={() => setFill(c)}
                    />
                  ))}
                </div>
              </div>

              <div className="xl-fc-row" style={{ marginTop: 8 }}>
                <span>Custom Color:</span>
                <input
                  type="color"
                  className="xl-fc-input"
                  style={{ width: 44, padding: 0, cursor: 'pointer' }}
                  value={fill || '#FFFFFF'}
                  onChange={(e) => setFill(e.target.value)}
                />
              </div>

              <div className="xl-fc-group" style={{ marginTop: 8 }}>
                <span className="xl-fc-group-title">Sample</span>
                <div
                  style={{
                    height: 40,
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
                  onChange={(e) => setLocked(e.target.checked)}
                />
                <span>Locked</span>
              </label>

              <label className="xl-fc-checkbox-label">
                <input
                  type="checkbox"
                  checked={hidden}
                  onChange={(e) => setHidden(e.target.checked)}
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

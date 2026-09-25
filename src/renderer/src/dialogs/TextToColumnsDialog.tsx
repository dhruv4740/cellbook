import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { a1, normRange } from '../model/types'
import {
  ColumnFormatType,
  convertTextToColumns,
  splitDelimitedLine,
  parseFixedLine
} from '../data/textToColumns'
import './TextToColumnsDialog.css'

export interface TextToColumnsDialogProps {
  store: WorkbookStore
  onClose: () => void
}

export function TextToColumnsDialog({ store, onClose }: TextToColumnsDialogProps) {
  const sheet = store.active
  const range = normRange(sheet.selection.ranges[0] || { r1: 0, c1: 0, r2: 0, c2: 0 })

  const [step, setStep] = useState<1 | 2 | 3>(1)
  const [mode, setMode] = useState<'delimited' | 'fixed'>('delimited')

  // Step 2 delimited settings
  const [tabDelim, setTabDelim] = useState(true)
  const [semiDelim, setSemiDelim] = useState(false)
  const [commaDelim, setCommaDelim] = useState(false)
  const [spaceDelim, setSpaceDelim] = useState(false)
  const [otherDelim, setOtherDelim] = useState(false)
  const [otherChar, setOtherChar] = useState('')
  const [consecutiveAsOne, setConsecutiveAsOne] = useState(false)
  const [textQualifier, setTextQualifier] = useState('"')

  // Step 2 fixed settings
  const [breakPositions, setBreakPositions] = useState<number[]>([10, 20])

  // Step 3 settings
  const [destination, setDestination] = useState<string>('$' + a1(range.r1, range.c1))
  const [columnFormats, setColumnFormats] = useState<ColumnFormatType[]>([])
  const [selectedCol, setSelectedCol] = useState(0)

  // Read preview raw lines (up to 15 rows)
  const previewLines: string[] = []
  for (let r = range.r1; r <= Math.min(range.r2, range.r1 + 14); r++) {
    previewLines.push(store.getRaw(sheet, r, range.c1))
  }

  // Active delimiters array
  const activeDelims: string[] = []
  if (tabDelim) activeDelims.push('\t')
  if (semiDelim) activeDelims.push(';')
  if (commaDelim) activeDelims.push(',')
  if (spaceDelim) activeDelims.push(' ')
  if (otherDelim && otherChar) activeDelims.push(otherChar)

  // Parse lines for preview
  const parsedPreviewRows = previewLines.map((line) => {
    if (mode === 'fixed') {
      return parseFixedLine(line, breakPositions)
    }
    return splitDelimitedLine(line, activeDelims, consecutiveAsOne, textQualifier)
  })

  const maxCols = Math.max(1, ...parsedPreviewRows.map((r) => r.length))

  const handleFinish = () => {
    convertTextToColumns(store, sheet, range, {
      mode,
      delimiters: {
        tab: tabDelim,
        semicolon: semiDelim,
        comma: commaDelim,
        space: spaceDelim,
        other: otherDelim,
        otherChar
      },
      consecutiveAsOne,
      textQualifier,
      breakPositions,
      columnFormats,
      destination
    })
    onClose()
  }

  const setColFmt = (fmt: ColumnFormatType) => {
    setColumnFormats((prev) => {
      const copy = [...prev]
      while (copy.length <= selectedCol) copy.push('general')
      copy[selectedCol] = fmt
      return copy
    })
  }

  const curColFormat = columnFormats[selectedCol] || 'general'

  return (
    <Dialog
      title={`Convert Text to Columns Wizard - Step ${step} of 3`}
      width={560}
      hideFooter
      onCancel={onClose}
    >
      <div className="xl-ttc-body">
        {step === 1 && (
          <div className="xl-ttc-step1">
            <p className="xl-ttc-desc">
              The Text Wizard has determined that your data is {mode}.
              <br />
              If this is correct, choose Next, or choose the data type that best describes your data.
            </p>

            <fieldset className="xl-ttc-fieldset">
              <legend>Original data type</legend>
              <label className="xl-ttc-radio-label">
                <input
                  type="radio"
                  name="ttcMode"
                  checked={mode === 'delimited'}
                  onChange={() => setMode('delimited')}
                />
                <div>
                  <strong>Delimited</strong> - Characters such as commas or tabs separate each field.
                </div>
              </label>
              <label className="xl-ttc-radio-label">
                <input
                  type="radio"
                  name="ttcMode"
                  checked={mode === 'fixed'}
                  onChange={() => setMode('fixed')}
                />
                <div>
                  <strong>Fixed width</strong> - Fields are aligned in columns with spaces between each field.
                </div>
              </label>
            </fieldset>

            <div className="xl-ttc-preview-section">
              <div className="xl-ttc-preview-title">Preview of selected data:</div>
              <div className="xl-ttc-preview-box">
                {previewLines.map((line, idx) => (
                  <div key={idx} className="xl-ttc-preview-line">
                    {line || '\u00A0'}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {step === 2 && (
          <div className="xl-ttc-step2">
            <p className="xl-ttc-desc">
              This screen lets you set the delimiters your data contains. You can see how your text is affected in the preview below.
            </p>

            <div className="xl-ttc-step2-controls">
              <fieldset className="xl-ttc-fieldset" style={{ flex: 1 }}>
                <legend>Delimiters</legend>
                <div className="xl-ttc-delim-grid">
                  <label>
                    <input
                      type="checkbox"
                      checked={tabDelim}
                      onChange={(e) => setTabDelim(e.target.checked)}
                    />
                    Tab
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={semiDelim}
                      onChange={(e) => setSemiDelim(e.target.checked)}
                    />
                    Semicolon
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={commaDelim}
                      onChange={(e) => setCommaDelim(e.target.checked)}
                    />
                    Comma
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={spaceDelim}
                      onChange={(e) => setSpaceDelim(e.target.checked)}
                    />
                    Space
                  </label>
                  <label className="xl-ttc-other-row">
                    <input
                      type="checkbox"
                      checked={otherDelim}
                      onChange={(e) => setOtherDelim(e.target.checked)}
                    />
                    Other:
                    <input
                      className="xl-ttc-input"
                      style={{ width: 30 }}
                      maxLength={1}
                      value={otherChar}
                      onChange={(e) => setOtherChar(e.target.value)}
                    />
                  </label>
                </div>
                <div style={{ marginTop: 8 }}>
                  <label>
                    <input
                      type="checkbox"
                      checked={consecutiveAsOne}
                      onChange={(e) => setConsecutiveAsOne(e.target.checked)}
                    />
                    Treat consecutive delimiters as one
                  </label>
                </div>
              </fieldset>

              <div className="xl-ttc-qualifier-col">
                <label>Text qualifier:</label>
                <select
                  className="xl-ttc-select"
                  value={textQualifier}
                  onChange={(e) => setTextQualifier(e.target.value)}
                >
                  <option value='"'>"</option>
                  <option value="'">'</option>
                  <option value="">&#123;none&#125;</option>
                </select>
              </div>
            </div>

            <div className="xl-ttc-preview-section">
              <div className="xl-ttc-preview-title">Data preview:</div>
              <div className="xl-ttc-table-preview">
                <table className="xl-ttc-table">
                  <tbody>
                    {parsedPreviewRows.map((row, rIdx) => (
                      <tr key={rIdx}>
                        {Array.from({ length: maxCols }).map((_, cIdx) => (
                          <td key={cIdx}>{row[cIdx] || '\u00A0'}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {step === 3 && (
          <div className="xl-ttc-step3">
            <p className="xl-ttc-desc">
              This screen lets you select each column and set the Data Format.
            </p>

            <div className="xl-ttc-step3-top">
              <fieldset className="xl-ttc-fieldset" style={{ flex: 1 }}>
                <legend>Column data format</legend>
                <label className="xl-ttc-radio-label">
                  <input
                    type="radio"
                    name="colFmt"
                    checked={curColFormat === 'general'}
                    onChange={() => setColFmt('general')}
                  />
                  <div>
                    <strong>General</strong> (converts numbers to numbers, dates to dates, rest to text)
                  </div>
                </label>
                <label className="xl-ttc-radio-label">
                  <input
                    type="radio"
                    name="colFmt"
                    checked={curColFormat === 'text'}
                    onChange={() => setColFmt('text')}
                  />
                  <span>Text</span>
                </label>
                <label className="xl-ttc-radio-label">
                  <input
                    type="radio"
                    name="colFmt"
                    checked={curColFormat === 'date'}
                    onChange={() => setColFmt('date')}
                  />
                  <span>Date (MDY)</span>
                </label>
                <label className="xl-ttc-radio-label">
                  <input
                    type="radio"
                    name="colFmt"
                    checked={curColFormat === 'skip'}
                    onChange={() => setColFmt('skip')}
                  />
                  <span>Do not import column (skip)</span>
                </label>
              </fieldset>

              <div className="xl-ttc-dest-col">
                <label>Destination:</label>
                <input
                  className="xl-ttc-input"
                  style={{ width: 100 }}
                  value={destination}
                  onChange={(e) => setDestination(e.target.value)}
                />
              </div>
            </div>

            <div className="xl-ttc-preview-section">
              <div className="xl-ttc-preview-title">Data preview:</div>
              <div className="xl-ttc-table-preview">
                <table className="xl-ttc-table">
                  <thead>
                    <tr>
                      {Array.from({ length: maxCols }).map((_, cIdx) => {
                        const fmt = columnFormats[cIdx] || 'general'
                        const isSel = cIdx === selectedCol
                        const label =
                          fmt === 'skip'
                            ? 'Skip'
                            : fmt.charAt(0).toUpperCase() + fmt.slice(1)
                        return (
                          <th
                            key={cIdx}
                            className={isSel ? 'selected' : ''}
                            onClick={() => setSelectedCol(cIdx)}
                          >
                            {label}
                          </th>
                        )
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {parsedPreviewRows.map((row, rIdx) => (
                      <tr key={rIdx}>
                        {Array.from({ length: maxCols }).map((_, cIdx) => (
                          <td
                            key={cIdx}
                            className={cIdx === selectedCol ? 'selected' : ''}
                            onClick={() => setSelectedCol(cIdx)}
                          >
                            {row[cIdx] || '\u00A0'}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        <div className="xl-ttc-footer">
          <button className="xl-ttc-btn" onClick={onClose}>
            Cancel
          </button>
          <div className="xl-ttc-nav-btns">
            <button
              className="xl-ttc-btn"
              disabled={step === 1}
              onClick={() => setStep((s) => (s > 1 ? ((s - 1) as any) : s))}
            >
              &lt; Back
            </button>
            {step < 3 ? (
              <button
                className="xl-ttc-btn primary"
                onClick={() => setStep((s) => (s < 3 ? ((s + 1) as any) : s))}
              >
                Next &gt;
              </button>
            ) : (
              <button className="xl-ttc-btn primary" onClick={handleFinish}>
                Finish
              </button>
            )}
          </div>
        </div>
      </div>
    </Dialog>
  )
}

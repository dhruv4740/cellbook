import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { SpecialType, SpecialSubtypes, getSpecialRanges } from './find'
import './GoToSpecialDialog.css'

export interface GoToSpecialDialogProps {
  store: WorkbookStore
  initialType?: SpecialType
  onClose: () => void
}

export function GoToSpecialDialog({
  store,
  initialType = 'formulas',
  onClose
}: GoToSpecialDialogProps) {
  const [selectedType, setSelectedType] = useState<SpecialType>(initialType)
  const [subtypes, setSubtypes] = useState<SpecialSubtypes>({
    numbers: true,
    text: true,
    logicals: true,
    errors: true
  })

  const [precedentsDirectOnly, setPrecedentsDirectOnly] = useState(true)
  const [validationAll, setValidationAll] = useState(true)
  const [message, setMessage] = useState<string | null>(null)

  const handleOk = () => {
    const ranges = getSpecialRanges(store, selectedType, subtypes)
    if (ranges && ranges.length > 0) {
      store.setSelection({
        ranges,
        active: { r: ranges[0].r1, c: ranges[0].c1 }
      })
      onClose()
    } else {
      setMessage('No cells were found.')
    }
  }

  const isFormulaOrConst = selectedType === 'formulas' || selectedType === 'constants'

  return (
    <Dialog title="Go To Special" width={380} onOk={handleOk} onCancel={onClose}>
      <div className="xl-gts-body">
        <div className="xl-gts-columns">
          {/* Left Column */}
          <div className="xl-gts-col">
            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'comments'}
                onChange={() => setSelectedType('comments')}
              />
              <span><u>C</u>omments</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'constants'}
                onChange={() => setSelectedType('constants')}
              />
              <span>O<u>x</u> Constants</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'formulas'}
                onChange={() => setSelectedType('formulas')}
              />
              <span><u>F</u>ormulas</span>
            </label>

            {/* Subtypes for Constants & Formulas */}
            <div className={`xl-gts-subgroup ${!isFormulaOrConst ? 'disabled' : ''}`}>
              <label className="xl-gts-check-label">
                <input
                  type="checkbox"
                  checked={subtypes.numbers}
                  disabled={!isFormulaOrConst}
                  onChange={(e) => setSubtypes({ ...subtypes, numbers: e.target.checked })}
                />
                <span><u>N</u>umbers</span>
              </label>
              <label className="xl-gts-check-label">
                <input
                  type="checkbox"
                  checked={subtypes.text}
                  disabled={!isFormulaOrConst}
                  onChange={(e) => setSubtypes({ ...subtypes, text: e.target.checked })}
                />
                <span><u>T</u>ext</span>
              </label>
              <label className="xl-gts-check-label">
                <input
                  type="checkbox"
                  checked={subtypes.logicals}
                  disabled={!isFormulaOrConst}
                  onChange={(e) => setSubtypes({ ...subtypes, logicals: e.target.checked })}
                />
                <span><u>L</u>ogicals</span>
              </label>
              <label className="xl-gts-check-label">
                <input
                  type="checkbox"
                  checked={subtypes.errors}
                  disabled={!isFormulaOrConst}
                  onChange={(e) => setSubtypes({ ...subtypes, errors: e.target.checked })}
                />
                <span><u>E</u>rrors</span>
              </label>
            </div>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'blanks'}
                onChange={() => setSelectedType('blanks')}
              />
              <span><u>B</u>lanks</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'currentRegion'}
                onChange={() => setSelectedType('currentRegion')}
              />
              <span>Current <u>r</u>egion</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'currentArray'}
                onChange={() => setSelectedType('currentArray')}
              />
              <span>Current <u>a</u>rray</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'objects'}
                onChange={() => setSelectedType('objects')}
              />
              <span><u>O</u>bjects</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'rowDifferences'}
                onChange={() => setSelectedType('rowDifferences')}
              />
              <span>Row di<u>f</u>ferences</span>
            </label>
          </div>

          {/* Right Column */}
          <div className="xl-gts-col">
            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'colDifferences'}
                onChange={() => setSelectedType('colDifferences')}
              />
              <span>Column differen<u>c</u>es</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'precedents'}
                onChange={() => setSelectedType('precedents')}
              />
              <span><u>P</u>recedents</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'dependents'}
                onChange={() => setSelectedType('dependents')}
              />
              <span><u>D</u>ependents</span>
            </label>

            <div
              className={`xl-gts-subgroup ${
                selectedType !== 'precedents' && selectedType !== 'dependents' ? 'disabled' : ''
              }`}
            >
              <label className="xl-gts-radio-label" style={{ fontSize: 11 }}>
                <input
                  type="radio"
                  name="predepScope"
                  checked={precedentsDirectOnly}
                  disabled={selectedType !== 'precedents' && selectedType !== 'dependents'}
                  onChange={() => setPrecedentsDirectOnly(true)}
                />
                <span>D<u>i</u>rect only</span>
              </label>
              <label className="xl-gts-radio-label" style={{ fontSize: 11 }}>
                <input
                  type="radio"
                  name="predepScope"
                  checked={!precedentsDirectOnly}
                  disabled={selectedType !== 'precedents' && selectedType !== 'dependents'}
                  onChange={() => setPrecedentsDirectOnly(false)}
                />
                <span>A<u>l</u>l levels</span>
              </label>
            </div>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'lastCell'}
                onChange={() => setSelectedType('lastCell')}
              />
              <span>La<u>s</u>t cell</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'visibleCells'}
                onChange={() => setSelectedType('visibleCells')}
              />
              <span><u>V</u>isible cells only</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'conditionalFormats'}
                onChange={() => setSelectedType('conditionalFormats')}
              />
              <span>Conditional forma<u>t</u>s</span>
            </label>

            <label className="xl-gts-radio-label">
              <input
                type="radio"
                name="specialType"
                checked={selectedType === 'dataValidation'}
                onChange={() => setSelectedType('dataValidation')}
              />
              <span>Data <u>v</u>alidation</span>
            </label>

            <div
              className={`xl-gts-subgroup ${
                selectedType !== 'dataValidation' ? 'disabled' : ''
              }`}
            >
              <label className="xl-gts-radio-label" style={{ fontSize: 11 }}>
                <input
                  type="radio"
                  name="validationScope"
                  checked={validationAll}
                  disabled={selectedType !== 'dataValidation'}
                  onChange={() => setValidationAll(true)}
                />
                <span><u>A</u>ll</span>
              </label>
              <label className="xl-gts-radio-label" style={{ fontSize: 11 }}>
                <input
                  type="radio"
                  name="validationScope"
                  checked={!validationAll}
                  disabled={selectedType !== 'dataValidation'}
                  onChange={() => setValidationAll(false)}
                />
                <span>Sa<u>m</u>e</span>
              </label>
            </div>
          </div>
        </div>

        {/* Excel-style Alert Box */}
        {message && (
          <div
            style={{
              position: 'fixed',
              inset: 0,
              background: 'rgba(0,0,0,0.2)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 1000
            }}
            onClick={() => setMessage(null)}
          >
            <div
              style={{
                background: '#fff',
                border: '1px solid #adadad',
                boxShadow: '0 4px 16px rgba(0,0,0,0.25)',
                width: 300,
                padding: 16
              }}
              onClick={(e) => e.stopPropagation()}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 16 }}>
                <span style={{ fontSize: 24, color: '#0078d7' }}>ℹ</span>
                <div style={{ fontSize: 12 }}>{message}</div>
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                <button
                  className="xl-btn primary"
                  autoFocus
                  onClick={() => setMessage(null)}
                >
                  OK
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </Dialog>
  )
}

import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import {
  DataValidationRule,
  DvAllowType,
  DvErrorStyle,
  DvOperator
} from '../model/types'
import {
  addDataValidation,
  clearDataValidations,
  getValidationForCell
} from '../data/dataValidation'
import './DataValidationDialog.css'

export interface DataValidationDialogProps {
  store: WorkbookStore
  onClose: () => void
}

export function DataValidationDialog({ store, onClose }: DataValidationDialogProps) {
  const sheet = store.active
  const ranges = sheet.selection.ranges
  const active = sheet.selection.active
  const existingRule = getValidationForCell(sheet, active.r, active.c)

  const [activeTab, setActiveTab] = useState<'settings' | 'inputMessage' | 'errorAlert'>('settings')

  // Settings
  const [allow, setAllow] = useState<DvAllowType>(existingRule?.allow || 'any')
  const [operator, setOperator] = useState<DvOperator>(existingRule?.operator || 'between')
  const [formula1, setFormula1] = useState(existingRule?.formula1 || '')
  const [formula2, setFormula2] = useState(existingRule?.formula2 || '')
  const [ignoreBlank, setIgnoreBlank] = useState(existingRule?.ignoreBlank ?? true)
  const [inCellDropdown, setInCellDropdown] = useState(existingRule?.inCellDropdown ?? true)

  // Input Message
  const [showInputMsg, setShowInputMsg] = useState(existingRule?.showInputMessage ?? true)
  const [inputTitle, setInputTitle] = useState(existingRule?.inputTitle || '')
  const [inputMessage, setInputMessage] = useState(existingRule?.inputMessage || '')

  // Error Alert
  const [showErrorAlert, setShowErrorAlert] = useState(existingRule?.showErrorAlert ?? true)
  const [errorStyle, setErrorStyle] = useState<DvErrorStyle>(existingRule?.errorStyle || 'stop')
  const [errorTitle, setErrorTitle] = useState(existingRule?.errorTitle || '')
  const [errorMessage, setErrorMessage] = useState(existingRule?.errorMessage || '')

  const handleClearAll = () => {
    clearDataValidations(store, sheet, ranges)
    setAllow('any')
    setFormula1('')
    setFormula2('')
    setInputTitle('')
    setInputMessage('')
    setErrorTitle('')
    setErrorMessage('')
    onClose()
  }

  const handleOk = () => {
    if (allow === 'any') {
      clearDataValidations(store, sheet, ranges)
    } else {
      addDataValidation(store, sheet, {
        ranges,
        allow,
        operator: allow === 'list' || allow === 'custom' ? undefined : operator,
        formula1,
        formula2: operator === 'between' || operator === 'notBetween' ? formula2 : undefined,
        ignoreBlank,
        inCellDropdown,
        showInputMessage: showInputMsg,
        inputTitle,
        inputMessage,
        showErrorAlert,
        errorStyle,
        errorTitle,
        errorMessage
      })
    }
    onClose()
  }

  const isNumericOrDate =
    allow === 'whole' ||
    allow === 'decimal' ||
    allow === 'date' ||
    allow === 'time' ||
    allow === 'textLength'

  return (
    <Dialog title="Data Validation" width={480} onOk={handleOk} onCancel={onClose}>
      <div className="xl-dv-body">
        <div className="xl-dv-tabs">
          <button
            className={`xl-dv-tab ${activeTab === 'settings' ? 'active' : ''}`}
            onClick={() => setActiveTab('settings')}
          >
            Settings
          </button>
          <button
            className={`xl-dv-tab ${activeTab === 'inputMessage' ? 'active' : ''}`}
            onClick={() => setActiveTab('inputMessage')}
          >
            Input Message
          </button>
          <button
            className={`xl-dv-tab ${activeTab === 'errorAlert' ? 'active' : ''}`}
            onClick={() => setActiveTab('errorAlert')}
          >
            Error Alert
          </button>
        </div>

        <div className="xl-dv-tab-content">
          {activeTab === 'settings' && (
            <div className="xl-dv-settings-pane">
              <div className="xl-dv-form-row">
                <div className="xl-dv-form-col" style={{ width: '48%' }}>
                  <label>Allow:</label>
                  <select
                    className="xl-dv-select"
                    value={allow}
                    onChange={(e) => setAllow(e.target.value as any)}
                  >
                    <option value="any">Any value</option>
                    <option value="whole">Whole number</option>
                    <option value="decimal">Decimal</option>
                    <option value="list">List</option>
                    <option value="date">Date</option>
                    <option value="time">Time</option>
                    <option value="textLength">Text length</option>
                    <option value="custom">Custom</option>
                  </select>
                </div>

                <div className="xl-dv-checkboxes-col" style={{ width: '48%' }}>
                  <label className="xl-dv-checkbox-label">
                    <input
                      type="checkbox"
                      checked={ignoreBlank}
                      onChange={(e) => setIgnoreBlank(e.target.checked)}
                    />
                    <span>Ignore blank</span>
                  </label>
                  {allow === 'list' && (
                    <label className="xl-dv-checkbox-label">
                      <input
                        type="checkbox"
                        checked={inCellDropdown}
                        onChange={(e) => setInCellDropdown(e.target.checked)}
                      />
                      <span>In-cell dropdown</span>
                    </label>
                  )}
                </div>
              </div>

              {isNumericOrDate && (
                <div className="xl-dv-form-col">
                  <label>Data:</label>
                  <select
                    className="xl-dv-select"
                    style={{ width: '48%' }}
                    value={operator}
                    onChange={(e) => setOperator(e.target.value as any)}
                  >
                    <option value="between">between</option>
                    <option value="notBetween">not between</option>
                    <option value="equal">equal to</option>
                    <option value="notEqual">not equal to</option>
                    <option value="greaterThan">greater than</option>
                    <option value="lessThan">less than</option>
                    <option value="greaterThanOrEqual">greater than or equal to</option>
                    <option value="lessThanOrEqual">less than or equal to</option>
                  </select>

                  <div className="xl-dv-form-row" style={{ marginTop: 8 }}>
                    <div className="xl-dv-form-col" style={{ flex: 1 }}>
                      <label>
                        {operator === 'between' || operator === 'notBetween'
                          ? 'Minimum:'
                          : 'Value:'}
                      </label>
                      <input
                        className="xl-dv-input"
                        value={formula1}
                        onChange={(e) => setFormula1(e.target.value)}
                      />
                    </div>

                    {(operator === 'between' || operator === 'notBetween') && (
                      <div className="xl-dv-form-col" style={{ flex: 1 }}>
                        <label>Maximum:</label>
                        <input
                          className="xl-dv-input"
                          value={formula2}
                          onChange={(e) => setFormula2(e.target.value)}
                        />
                      </div>
                    )}
                  </div>
                </div>
              )}

              {allow === 'list' && (
                <div className="xl-dv-form-col" style={{ marginTop: 8 }}>
                  <label>Source:</label>
                  <input
                    className="xl-dv-input"
                    placeholder="e.g. Yes, No, Maybe or =A1:A5"
                    value={formula1}
                    onChange={(e) => setFormula1(e.target.value)}
                  />
                </div>
              )}

              {allow === 'custom' && (
                <div className="xl-dv-form-col" style={{ marginTop: 8 }}>
                  <label>Formula:</label>
                  <input
                    className="xl-dv-input"
                    placeholder="=A1>0"
                    value={formula1}
                    onChange={(e) => setFormula1(e.target.value)}
                  />
                </div>
              )}
            </div>
          )}

          {activeTab === 'inputMessage' && (
            <div className="xl-dv-input-msg-pane">
              <label className="xl-dv-checkbox-label">
                <input
                  type="checkbox"
                  checked={showInputMsg}
                  onChange={(e) => setShowInputMsg(e.target.checked)}
                />
                <span>Show input message when cell is selected</span>
              </label>

              <div className="xl-dv-form-col" style={{ marginTop: 12 }}>
                <label>Title:</label>
                <input
                  className="xl-dv-input"
                  disabled={!showInputMsg}
                  value={inputTitle}
                  onChange={(e) => setInputTitle(e.target.value)}
                />
              </div>

              <div className="xl-dv-form-col" style={{ marginTop: 8 }}>
                <label>Input message:</label>
                <textarea
                  className="xl-dv-textarea"
                  rows={4}
                  disabled={!showInputMsg}
                  value={inputMessage}
                  onChange={(e) => setInputMessage(e.target.value)}
                />
              </div>
            </div>
          )}

          {activeTab === 'errorAlert' && (
            <div className="xl-dv-error-alert-pane">
              <label className="xl-dv-checkbox-label">
                <input
                  type="checkbox"
                  checked={showErrorAlert}
                  onChange={(e) => setShowErrorAlert(e.target.checked)}
                />
                <span>Show error alert after invalid data is entered</span>
              </label>

              <div className="xl-dv-form-row" style={{ marginTop: 12 }}>
                <div className="xl-dv-form-col" style={{ width: 120 }}>
                  <label>Style:</label>
                  <select
                    className="xl-dv-select"
                    disabled={!showErrorAlert}
                    value={errorStyle}
                    onChange={(e) => setErrorStyle(e.target.value as any)}
                  >
                    <option value="stop">Stop</option>
                    <option value="warning">Warning</option>
                    <option value="information">Information</option>
                  </select>
                </div>

                <div className="xl-dv-form-col" style={{ flex: 1 }}>
                  <label>Title:</label>
                  <input
                    className="xl-dv-input"
                    disabled={!showErrorAlert}
                    value={errorTitle}
                    onChange={(e) => setErrorTitle(e.target.value)}
                  />
                </div>
              </div>

              <div className="xl-dv-form-col" style={{ marginTop: 8 }}>
                <label>Error message:</label>
                <textarea
                  className="xl-dv-textarea"
                  rows={4}
                  disabled={!showErrorAlert}
                  value={errorMessage}
                  onChange={(e) => setErrorMessage(e.target.value)}
                />
              </div>
            </div>
          )}
        </div>

        <div className="xl-dv-footer">
          <button className="xl-dv-clear-btn" onClick={handleClearAll}>
            Clear All
          </button>
        </div>
      </div>
    </Dialog>
  )
}

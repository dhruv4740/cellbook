import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { CellStyle, CfDateOccurring, CfOperator, CfRule, CfRuleType } from '../model/types'
import { addCfRule } from '../format/conditional'
import './CfNewRuleDialog.css'

export interface CfNewRuleDialogProps {
  store: WorkbookStore
  editRule?: CfRule
  onClose: () => void
}

const RULE_TYPES = [
  'Format all cells based on their values',
  'Format only cells that contain',
  'Format only top or bottom ranked values',
  'Format only values that are above or below average',
  'Format only unique or duplicate values',
  'Use a formula to determine which cells to format'
]

export function CfNewRuleDialog({ store, editRule, onClose }: CfNewRuleDialogProps) {
  const sheet = store.active
  const ranges = editRule ? editRule.ranges : sheet.selection.ranges

  // Type selection index (0..5)
  const initialTypeIdx = () => {
    if (!editRule) return 1
    if (editRule.type === 'dataBar' || editRule.type === 'colorScale' || editRule.type === 'iconSet') return 0
    if (editRule.type === 'cellIs' || editRule.type === 'textContains' || editRule.type === 'dateOccurring') return 1
    if (editRule.type === 'top10') return 2
    if (editRule.type === 'aboveAverage') return 3
    if (editRule.type === 'duplicateValues' || editRule.type === 'uniqueValues') return 4
    if (editRule.type === 'expression') return 5
    return 1
  }

  const [typeIdx, setTypeIdx] = useState<number>(initialTypeIdx)

  // Type 0: Format all cells based on their values
  const [styleKind, setStyleKind] = useState<'2color' | '3color' | 'dataBar' | 'iconSet'>(() => {
    if (editRule?.type === 'dataBar') return 'dataBar'
    if (editRule?.type === 'iconSet') return 'iconSet'
    if (editRule?.type === 'colorScale') {
      return editRule.colorScale?.midColor ? '3color' : '2color'
    }
    return '2color'
  })
  const [scaleMinColor, setScaleMinColor] = useState(editRule?.colorScale?.minColor || '#F8696B')
  const [scaleMidColor, setScaleMidColor] = useState(editRule?.colorScale?.midColor || '#FFEB84')
  const [scaleMaxColor, setScaleMaxColor] = useState(editRule?.colorScale?.maxColor || '#63BE7B')
  const [barColor, setBarColor] = useState(editRule?.dataBar?.color || '#638EC6')
  const [barGradient, setBarGradient] = useState(editRule?.dataBar?.gradient ?? true)
  const [iconSetName, setIconSetName] = useState(editRule?.iconSet?.name || '3TrafficLights1')

  // Type 1: Format only cells that contain
  const [containKind, setContainKind] = useState<'cellValue' | 'specificText' | 'datesOccurring'>(() => {
    if (editRule?.type === 'textContains') return 'specificText'
    if (editRule?.type === 'dateOccurring') return 'datesOccurring'
    return 'cellValue'
  })
  const [operator, setOperator] = useState<CfOperator>(editRule?.operator || 'between')
  const [val1, setVal1] = useState(editRule?.value || '0')
  const [val2, setVal2] = useState(editRule?.value2 || '100')
  const [textVal, setTextVal] = useState(editRule?.text || '')
  const [dateVal, setDateVal] = useState<CfDateOccurring>(editRule?.date || 'yesterday')

  // Type 2: Top / Bottom
  const [topBottom, setTopBottom] = useState<'top' | 'bottom'>(editRule?.bottom ? 'bottom' : 'top')
  const [rankVal, setRankVal] = useState(String(editRule?.rank || 10))
  const [percentVal, setPercentVal] = useState(Boolean(editRule?.percent))

  // Type 3: Above / Below average
  const [aboveBelow, setAboveBelow] = useState<'above' | 'below' | 'equalOrAbove' | 'equalOrBelow'>(() => {
    if (editRule?.equalOrAbove) return 'equalOrAbove'
    if (editRule?.above === false) return 'below'
    return 'above'
  })

  // Type 4: Duplicate / Unique
  const [dupOrUnique, setDupOrUnique] = useState<'duplicate' | 'unique'>(
    editRule?.type === 'uniqueValues' ? 'unique' : 'duplicate'
  )

  // Type 5: Expression formula
  const [formulaVal, setFormulaVal] = useState(editRule?.formula || '=A1>0')

  // Format styling
  const [fontColor, setFontColor] = useState(editRule?.format?.color || '#9C0006')
  const [fillColor, setFillColor] = useState(editRule?.format?.fill || '#FFC7CE')
  const [isBold, setIsBold] = useState(Boolean(editRule?.format?.bold))
  const [isItalic, setIsItalic] = useState(Boolean(editRule?.format?.italic))
  const [hasBorder, setHasBorder] = useState(Boolean(editRule?.format?.borders))

  const handleOk = () => {
    let finalRule: Omit<CfRule, 'id' | 'priority'>

    const resolvedFormat: Partial<CellStyle> = {
      color: fontColor || undefined,
      fill: fillColor || undefined,
      bold: isBold || undefined,
      italic: isItalic || undefined,
      borders: hasBorder
        ? {
            top: { style: 'thin', color: '#000000' },
            bottom: { style: 'thin', color: '#000000' },
            left: { style: 'thin', color: '#000000' },
            right: { style: 'thin', color: '#000000' }
          }
        : undefined
    }

    if (typeIdx === 0) {
      if (styleKind === '2color') {
        finalRule = {
          ranges,
          type: 'colorScale',
          colorScale: { minColor: scaleMinColor, maxColor: scaleMaxColor }
        }
      } else if (styleKind === '3color') {
        finalRule = {
          ranges,
          type: 'colorScale',
          colorScale: { minColor: scaleMinColor, midColor: scaleMidColor, maxColor: scaleMaxColor }
        }
      } else if (styleKind === 'dataBar') {
        finalRule = {
          ranges,
          type: 'dataBar',
          dataBar: { color: barColor, gradient: barGradient }
        }
      } else {
        finalRule = {
          ranges,
          type: 'iconSet',
          iconSet: { name: iconSetName }
        }
      }
    } else if (typeIdx === 1) {
      if (containKind === 'cellValue') {
        finalRule = {
          ranges,
          type: 'cellIs',
          operator,
          value: val1,
          value2: operator === 'between' || operator === 'notBetween' ? val2 : undefined,
          format: resolvedFormat
        }
      } else if (containKind === 'specificText') {
        finalRule = {
          ranges,
          type: 'textContains',
          text: textVal,
          format: resolvedFormat
        }
      } else {
        finalRule = {
          ranges,
          type: 'dateOccurring',
          date: dateVal,
          format: resolvedFormat
        }
      }
    } else if (typeIdx === 2) {
      finalRule = {
        ranges,
        type: 'top10',
        bottom: topBottom === 'bottom',
        rank: +rankVal || 10,
        percent: percentVal,
        format: resolvedFormat
      }
    } else if (typeIdx === 3) {
      finalRule = {
        ranges,
        type: 'aboveAverage',
        above: aboveBelow === 'above' || aboveBelow === 'equalOrAbove',
        equalOrAbove: aboveBelow === 'equalOrAbove',
        format: resolvedFormat
      }
    } else if (typeIdx === 4) {
      finalRule = {
        ranges,
        type: dupOrUnique === 'duplicate' ? 'duplicateValues' : 'uniqueValues',
        format: resolvedFormat
      }
    } else {
      finalRule = {
        ranges,
        type: 'expression',
        formula: formulaVal.startsWith('=') ? formulaVal : '=' + formulaVal,
        format: resolvedFormat
      }
    }

    if (editRule) {
      store.updateLayout(sheet, 'Edit Conditional Formatting Rule', (l) => {
        l.conditionalFormats = (l.conditionalFormats || []).map((r) =>
          r.id === editRule.id ? { ...finalRule, id: editRule.id, priority: editRule.priority } : r
        )
      })
    } else {
      addCfRule(store, sheet, finalRule)
    }

    onClose()
  }

  return (
    <Dialog
      title={editRule ? 'Edit Formatting Rule' : 'New Formatting Rule'}
      width={520}
      onOk={handleOk}
      onCancel={onClose}
    >
      <div className="xl-cf-new-body">
        <label className="xl-cf-section-label">Select a Rule Type:</label>
        <div className="xl-cf-type-listbox" role="listbox">
          {RULE_TYPES.map((label, idx) => (
            <div
              key={idx}
              className={`xl-cf-type-item ${typeIdx === idx ? 'selected' : ''}`}
              onClick={() => setTypeIdx(idx)}
            >
              {label}
            </div>
          ))}
        </div>

        <label className="xl-cf-section-label">Edit the Rule Description:</label>
        <div className="xl-cf-desc-box">
          {typeIdx === 0 && (
            <div className="xl-cf-form-col">
              <div className="xl-cf-form-row">
                <span>Format Style:</span>
                <select
                  className="xl-cf-select"
                  value={styleKind}
                  onChange={(e) => setStyleKind(e.target.value as any)}
                >
                  <option value="2color">2-Color Scale</option>
                  <option value="3color">3-Color Scale</option>
                  <option value="dataBar">Data Bar</option>
                  <option value="iconSet">Icon Sets</option>
                </select>
              </div>

              {styleKind === '2color' && (
                <div className="xl-cf-color-scale-row">
                  <div>
                    <label>Minimum:</label>
                    <input
                      type="color"
                      className="xl-cf-color-picker"
                      value={scaleMinColor}
                      onChange={(e) => setScaleMinColor(e.target.value)}
                    />
                  </div>
                  <div>
                    <label>Maximum:</label>
                    <input
                      type="color"
                      className="xl-cf-color-picker"
                      value={scaleMaxColor}
                      onChange={(e) => setScaleMaxColor(e.target.value)}
                    />
                  </div>
                </div>
              )}

              {styleKind === '3color' && (
                <div className="xl-cf-color-scale-row">
                  <div>
                    <label>Minimum:</label>
                    <input
                      type="color"
                      className="xl-cf-color-picker"
                      value={scaleMinColor}
                      onChange={(e) => setScaleMinColor(e.target.value)}
                    />
                  </div>
                  <div>
                    <label>Midpoint:</label>
                    <input
                      type="color"
                      className="xl-cf-color-picker"
                      value={scaleMidColor}
                      onChange={(e) => setScaleMidColor(e.target.value)}
                    />
                  </div>
                  <div>
                    <label>Maximum:</label>
                    <input
                      type="color"
                      className="xl-cf-color-picker"
                      value={scaleMaxColor}
                      onChange={(e) => setScaleMaxColor(e.target.value)}
                    />
                  </div>
                </div>
              )}

              {styleKind === 'dataBar' && (
                <div className="xl-cf-form-row">
                  <span>Bar Color:</span>
                  <input
                    type="color"
                    className="xl-cf-color-picker"
                    value={barColor}
                    onChange={(e) => setBarColor(e.target.value)}
                  />
                  <label className="xl-cf-checkbox-label">
                    <input
                      type="checkbox"
                      checked={barGradient}
                      onChange={(e) => setBarGradient(e.target.checked)}
                    />
                    Gradient fill
                  </label>
                </div>
              )}

              {styleKind === 'iconSet' && (
                <div className="xl-cf-form-row">
                  <span>Icon Style:</span>
                  <select
                    className="xl-cf-select"
                    value={iconSetName}
                    onChange={(e) => setIconSetName(e.target.value)}
                  >
                    <option value="3TrafficLights1">3 Traffic Lights (Unrimmed)</option>
                    <option value="3TrafficLights2">3 Traffic Lights (Rimmed)</option>
                    <option value="3Arrows">3 Arrows (Colored)</option>
                    <option value="3Symbols">3 Symbols (Circled)</option>
                  </select>
                </div>
              )}
            </div>
          )}

          {typeIdx === 1 && (
            <div className="xl-cf-form-col">
              <div className="xl-cf-form-row">
                <select
                  className="xl-cf-select"
                  value={containKind}
                  onChange={(e) => setContainKind(e.target.value as any)}
                >
                  <option value="cellValue">Cell Value</option>
                  <option value="specificText">Specific Text</option>
                  <option value="datesOccurring">Dates Occurring</option>
                </select>

                {containKind === 'cellValue' && (
                  <>
                    <select
                      className="xl-cf-select"
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
                    <input
                      className="xl-cf-input"
                      style={{ width: 60 }}
                      value={val1}
                      onChange={(e) => setVal1(e.target.value)}
                    />
                    {(operator === 'between' || operator === 'notBetween') && (
                      <>
                        <span>and</span>
                        <input
                          className="xl-cf-input"
                          style={{ width: 60 }}
                          value={val2}
                          onChange={(e) => setVal2(e.target.value)}
                        />
                      </>
                    )}
                  </>
                )}

                {containKind === 'specificText' && (
                  <>
                    <span>containing</span>
                    <input
                      className="xl-cf-input"
                      style={{ flex: 1 }}
                      value={textVal}
                      onChange={(e) => setTextVal(e.target.value)}
                    />
                  </>
                )}

                {containKind === 'datesOccurring' && (
                  <select
                    className="xl-cf-select"
                    value={dateVal}
                    onChange={(e) => setDateVal(e.target.value as any)}
                  >
                    <option value="yesterday">Yesterday</option>
                    <option value="today">Today</option>
                    <option value="tomorrow">Tomorrow</option>
                    <option value="inTheLast7Days">In the last 7 days</option>
                    <option value="lastWeek">Last week</option>
                    <option value="thisWeek">This week</option>
                    <option value="nextWeek">Next week</option>
                    <option value="lastMonth">Last month</option>
                    <option value="thisMonth">This month</option>
                    <option value="nextMonth">Next month</option>
                  </select>
                )}
              </div>
            </div>
          )}

          {typeIdx === 2 && (
            <div className="xl-cf-form-row">
              <span>Format values that rank in the:</span>
              <select
                className="xl-cf-select"
                value={topBottom}
                onChange={(e) => setTopBottom(e.target.value as any)}
              >
                <option value="top">Top</option>
                <option value="bottom">Bottom</option>
              </select>
              <input
                className="xl-cf-input"
                style={{ width: 50 }}
                type="number"
                min="1"
                value={rankVal}
                onChange={(e) => setRankVal(e.target.value)}
              />
              <label className="xl-cf-checkbox-label">
                <input
                  type="checkbox"
                  checked={percentVal}
                  onChange={(e) => setPercentVal(e.target.checked)}
                />
                % of the selected range
              </label>
            </div>
          )}

          {typeIdx === 3 && (
            <div className="xl-cf-form-row">
              <span>Format values that are:</span>
              <select
                className="xl-cf-select"
                value={aboveBelow}
                onChange={(e) => setAboveBelow(e.target.value as any)}
              >
                <option value="above">above</option>
                <option value="below">below</option>
                <option value="equalOrAbove">equal or above</option>
                <option value="equalOrBelow">equal or below</option>
              </select>
              <span>the average for the selected range</span>
            </div>
          )}

          {typeIdx === 4 && (
            <div className="xl-cf-form-row">
              <span>Format all:</span>
              <select
                className="xl-cf-select"
                value={dupOrUnique}
                onChange={(e) => setDupOrUnique(e.target.value as any)}
              >
                <option value="duplicate">duplicate</option>
                <option value="unique">unique</option>
              </select>
              <span>values in the selected range</span>
            </div>
          )}

          {typeIdx === 5 && (
            <div className="xl-cf-form-col">
              <span>Format values where this formula is true:</span>
              <input
                className="xl-cf-input"
                style={{ width: '100%' }}
                value={formulaVal}
                placeholder="=A1>0"
                onChange={(e) => setFormulaVal(e.target.value)}
              />
            </div>
          )}

          {typeIdx !== 0 && (
            <div className="xl-cf-format-preview-section">
              <div className="xl-cf-preview-wrapper">
                <span>Preview:</span>
                <div
                  className="xl-cf-preview-box"
                  style={{
                    backgroundColor: fillColor || '#ffffff',
                    color: fontColor || '#000000',
                    fontWeight: isBold ? 'bold' : 'normal',
                    fontStyle: isItalic ? 'italic' : 'normal',
                    border: hasBorder ? '1px solid #000000' : '1px solid #adadad'
                  }}
                >
                  AaBbCcYyZz
                </div>
              </div>

              <div className="xl-cf-format-controls">
                <label className="xl-cf-format-item">
                  <span>Fill:</span>
                  <input
                    type="color"
                    className="xl-cf-color-picker"
                    value={fillColor}
                    onChange={(e) => setFillColor(e.target.value)}
                  />
                </label>
                <label className="xl-cf-format-item">
                  <span>Font:</span>
                  <input
                    type="color"
                    className="xl-cf-color-picker"
                    value={fontColor}
                    onChange={(e) => setFontColor(e.target.value)}
                  />
                </label>
                <label className="xl-cf-checkbox-label">
                  <input
                    type="checkbox"
                    checked={isBold}
                    onChange={(e) => setIsBold(e.target.checked)}
                  />
                  Bold
                </label>
                <label className="xl-cf-checkbox-label">
                  <input
                    type="checkbox"
                    checked={isItalic}
                    onChange={(e) => setIsItalic(e.target.checked)}
                  />
                  Italic
                </label>
                <label className="xl-cf-checkbox-label">
                  <input
                    type="checkbox"
                    checked={hasBorder}
                    onChange={(e) => setHasBorder(e.target.checked)}
                  />
                  Border
                </label>
              </div>
            </div>
          )}
        </div>
      </div>
    </Dialog>
  )
}

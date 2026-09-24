import { useState, useEffect, useRef } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { CellStyle, CfDateOccurring, CfOperator, CfRule, CfRuleType } from '../model/types'
import { addCfRule } from '../format/conditional'
import './CfQuickDialog.css'

export type CfQuickKind =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'equal'
  | 'textContains'
  | 'dateOccurring'
  | 'duplicateValues'
  | 'top10Items'
  | 'top10Percent'
  | 'bottom10Items'
  | 'bottom10Percent'
  | 'aboveAverage'
  | 'belowAverage'

export interface CfQuickDialogProps {
  kind: CfQuickKind
  store: WorkbookStore
  onClose: () => void
}

const PRESETS: { label: string; format: Partial<CellStyle> }[] = [
  { label: 'Light Red Fill with Dark Red Text', format: { fill: '#FFC7CE', color: '#9C0006' } },
  { label: 'Yellow Fill with Dark Yellow Text', format: { fill: '#FFEB9C', color: '#9C6500' } },
  { label: 'Green Fill with Dark Green Text', format: { fill: '#C6EFCE', color: '#006100' } },
  { label: 'Light Red Fill', format: { fill: '#FFC7CE' } },
  { label: 'Red Text', format: { color: '#9C0006' } },
  {
    label: 'Red Border',
    format: {
      borders: {
        top: { style: 'thin', color: '#FF0000' },
        bottom: { style: 'thin', color: '#FF0000' },
        left: { style: 'thin', color: '#FF0000' },
        right: { style: 'thin', color: '#FF0000' }
      }
    }
  },
  { label: 'Custom Format...', format: { fill: '#D9E1F2', color: '#002060' } }
]

export function CfQuickDialog({ kind, store, onClose }: CfQuickDialogProps) {
  const sheet = store.active
  const ranges = sheet.selection.ranges

  // Value states
  const [val1, setVal1] = useState('10')
  const [val2, setVal2] = useState('20')
  const [dateOption, setDateOption] = useState<CfDateOccurring>('yesterday')
  const [dupOption, setDupOption] = useState<'duplicate' | 'unique'>('duplicate')
  const [presetIndex, setPresetIndex] = useState(0)
  const [customFill, setCustomFill] = useState('#D9E1F2')
  const [customColor, setCustomColor] = useState('#002060')

  // Save initial rules for live preview rollback
  const initialRulesRef = useRef<CfRule[]>(sheet.layout.conditionalFormats || [])
  const committedRef = useRef(false)

  let title = 'Greater Than'
  let promptText = 'Format cells that are GREATER THAN:'

  if (kind === 'lessThan') {
    title = 'Less Than'
    promptText = 'Format cells that are LESS THAN:'
  } else if (kind === 'between') {
    title = 'Between'
    promptText = 'Format cells that are BETWEEN:'
  } else if (kind === 'equal') {
    title = 'Equal To'
    promptText = 'Format cells that are EQUAL TO:'
  } else if (kind === 'textContains') {
    title = 'Text That Contains'
    promptText = 'Format cells that contain the text:'
  } else if (kind === 'dateOccurring') {
    title = 'A Date Occurring'
    promptText = 'Format cells that contain a date occurring:'
  } else if (kind === 'duplicateValues') {
    title = 'Duplicate Values'
    promptText = 'Format cells that contain:'
  } else if (kind === 'top10Items') {
    title = 'Top 10 Items'
    promptText = 'Format cells that rank in the TOP:'
  } else if (kind === 'top10Percent') {
    title = 'Top 10%'
    promptText = 'Format cells that rank in the TOP:'
  } else if (kind === 'bottom10Items') {
    title = 'Bottom 10 Items'
    promptText = 'Format cells that rank in the BOTTOM:'
  } else if (kind === 'bottom10Percent') {
    title = 'Bottom 10%'
    promptText = 'Format cells that rank in the BOTTOM:'
  } else if (kind === 'aboveAverage') {
    title = 'Above Average'
    promptText = 'Format values that are ABOVE the average for the selected range:'
  } else if (kind === 'belowAverage') {
    title = 'Below Average'
    promptText = 'Format values that are BELOW the average for the selected range:'
  }

  const buildRule = (): Omit<CfRule, 'id' | 'priority'> => {
    let chosenFormat = PRESETS[presetIndex].format
    if (presetIndex === 6) {
      chosenFormat = { fill: customFill, color: customColor }
    }

    let type: CfRuleType = 'cellIs'
    let operator: CfOperator | undefined
    let value: string | undefined = val1
    let value2: string | undefined
    let text: string | undefined
    let date: CfDateOccurring | undefined
    let rank: number | undefined
    let percent = false
    let bottom = false
    let above: boolean | undefined

    if (kind === 'greaterThan') {
      type = 'cellIs'
      operator = 'greaterThan'
    } else if (kind === 'lessThan') {
      type = 'cellIs'
      operator = 'lessThan'
    } else if (kind === 'between') {
      type = 'cellIs'
      operator = 'between'
      value2 = val2
    } else if (kind === 'equal') {
      type = 'cellIs'
      operator = 'equal'
    } else if (kind === 'textContains') {
      type = 'textContains'
      text = val1
      value = undefined
    } else if (kind === 'dateOccurring') {
      type = 'dateOccurring'
      date = dateOption
      value = undefined
    } else if (kind === 'duplicateValues') {
      type = dupOption === 'duplicate' ? 'duplicateValues' : 'uniqueValues'
      value = undefined
    } else if (kind === 'top10Items') {
      type = 'top10'
      rank = +val1 || 10
    } else if (kind === 'top10Percent') {
      type = 'top10'
      rank = +val1 || 10
      percent = true
    } else if (kind === 'bottom10Items') {
      type = 'top10'
      rank = +val1 || 10
      bottom = true
    } else if (kind === 'bottom10Percent') {
      type = 'top10'
      rank = +val1 || 10
      percent = true
      bottom = true
    } else if (kind === 'aboveAverage') {
      type = 'aboveAverage'
      above = true
      value = undefined
    } else if (kind === 'belowAverage') {
      type = 'aboveAverage'
      above = false
      value = undefined
    }

    return {
      ranges,
      type,
      operator,
      value,
      value2,
      text,
      date,
      percent,
      bottom,
      rank,
      above,
      format: chosenFormat
    }
  }

  // Live preview effect
  useEffect(() => {
    const previewRule: CfRule = {
      ...buildRule(),
      id: '__preview_cf__',
      priority: 0
    }

    sheet.layout.conditionalFormats = [
      previewRule,
      ...initialRulesRef.current.filter((r) => r.id !== '__preview_cf__')
    ]
    store.notify()

    return () => {
      if (!committedRef.current) {
        sheet.layout.conditionalFormats = initialRulesRef.current
        store.notify()
      }
    }
  }, [val1, val2, dateOption, dupOption, presetIndex, customFill, customColor])

  const handleOk = () => {
    committedRef.current = true
    sheet.layout.conditionalFormats = initialRulesRef.current
    addCfRule(store, sheet, buildRule())
    onClose()
  }

  const handleCancel = () => {
    sheet.layout.conditionalFormats = initialRulesRef.current
    store.notify()
    onClose()
  }

  return (
    <Dialog title={title} width={430} onOk={handleOk} onCancel={handleCancel}>
      <div className="xl-cf-quick-body">
        <div className="xl-cf-quick-label">{promptText}</div>
        <div className="xl-cf-quick-row">
          {kind === 'between' ? (
            <>
              <input
                className="xl-cf-quick-input"
                style={{ width: 70 }}
                value={val1}
                autoFocus
                onChange={(e) => setVal1(e.target.value)}
              />
              <span>and</span>
              <input
                className="xl-cf-quick-input"
                style={{ width: 70 }}
                value={val2}
                onChange={(e) => setVal2(e.target.value)}
              />
            </>
          ) : kind === 'dateOccurring' ? (
            <select
              className="xl-cf-quick-select"
              style={{ width: 140 }}
              value={dateOption}
              onChange={(e) => setDateOption(e.target.value as any)}
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
          ) : kind === 'duplicateValues' ? (
            <select
              className="xl-cf-quick-select"
              style={{ width: 110 }}
              value={dupOption}
              onChange={(e) => setDupOption(e.target.value as any)}
            >
              <option value="duplicate">duplicate</option>
              <option value="unique">unique</option>
            </select>
          ) : kind !== 'aboveAverage' && kind !== 'belowAverage' ? (
            <input
              className="xl-cf-quick-input"
              style={{ width: 90 }}
              value={val1}
              autoFocus
              onChange={(e) => setVal1(e.target.value)}
            />
          ) : null}

          <span>with</span>

          <select
            className="xl-cf-quick-select"
            style={{ flex: 1 }}
            value={presetIndex}
            onChange={(e) => setPresetIndex(+e.target.value)}
          >
            {PRESETS.map((p, i) => (
              <option key={i} value={i}>
                {p.label}
              </option>
            ))}
          </select>
        </div>

        {presetIndex === 6 && (
          <div className="xl-cf-quick-custom-row">
            <span>Custom Fill:</span>
            <input
              type="color"
              value={customFill}
              onChange={(e) => setCustomFill(e.target.value)}
            />
            <span>Custom Font:</span>
            <input
              type="color"
              value={customColor}
              onChange={(e) => setCustomColor(e.target.value)}
            />
          </div>
        )}
      </div>
    </Dialog>
  )
}

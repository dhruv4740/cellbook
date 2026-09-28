import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { WorkbookStore } from '../model/store'
import { CfRule, parseA1, rangeA1, normRange } from '../model/types'
import './CfManageDialog.css'

export interface CfManageDialogProps {
  store: WorkbookStore
  onClose: () => void
  onNewRule: () => void
  onEditRule: (rule: CfRule) => void
}

function ruleDescription(rule: CfRule): string {
  switch (rule.type) {
    case 'cellIs':
      if (rule.operator === 'between' || rule.operator === 'notBetween') {
        return `Cell Value ${rule.operator === 'between' ? 'between' : 'not between'} ${rule.value} and ${rule.value2}`
      }
      return `Cell Value ${rule.operator || 'equal'} ${rule.value ?? ''}`
    case 'textContains':
      return `Cell Value contains "${rule.text ?? ''}"`
    case 'dateOccurring':
      return `Date occurring: ${rule.date ?? ''}`
    case 'duplicateValues':
      return 'Duplicate Values'
    case 'uniqueValues':
      return 'Unique Values'
    case 'top10':
      return `${rule.bottom ? 'Bottom' : 'Top'} ${rule.rank || 10}${rule.percent ? '%' : ''}`
    case 'aboveAverage':
      return `${rule.above ? 'Above' : 'Below'} Average`
    case 'expression':
      return `Formula: ${rule.formula ?? ''}`
    case 'dataBar':
      return 'Data Bar'
    case 'colorScale':
      return 'Graded Color Scale'
    case 'iconSet':
      return `Icon Set: ${rule.iconSet?.name ?? ''}`
    default:
      return 'Conditional Rule'
  }
}

function rangesToString(rule: CfRule): string {
  return rule.ranges.map((rg) => '=' + rangeA1(rg)).join(', ')
}

export function CfManageDialog({ store, onClose, onNewRule, onEditRule }: CfManageDialogProps) {
  const sheet = store.active
  const [scope, setScope] = useState<'selection' | 'sheet'>('selection')
  const [rules, setRules] = useState<CfRule[]>(() => {
    return (sheet.layout.conditionalFormats || []).slice().sort((a, b) => a.priority - b.priority)
  })
  const [selectedId, setSelectedId] = useState<string | null>(() => rules[0]?.id || null)

  const selRanges = sheet.selection.ranges
  const visibleRules = rules.filter((r) => {
    if (scope === 'sheet') return true
    return r.ranges.some((rg) => {
      const nr = normRange(rg)
      return selRanges.some((sel) => {
        const ns = normRange(sel)
        return !(nr.r2 < ns.r1 || nr.r1 > ns.r2 || nr.c2 < ns.c1 || nr.c1 > ns.c2)
      })
    })
  })

  const selectedRule = rules.find((r) => r.id === selectedId)

  const handleMoveUp = () => {
    if (!selectedId) return
    const idx = rules.findIndex((r) => r.id === selectedId)
    if (idx <= 0) return
    const copy = [...rules]
    const temp = copy[idx]
    copy[idx] = copy[idx - 1]
    copy[idx - 1] = temp
    copy.forEach((r, i) => (r.priority = i + 1))
    setRules(copy)
  }

  const handleMoveDown = () => {
    if (!selectedId) return
    const idx = rules.findIndex((r) => r.id === selectedId)
    if (idx < 0 || idx >= rules.length - 1) return
    const copy = [...rules]
    const temp = copy[idx]
    copy[idx] = copy[idx + 1]
    copy[idx + 1] = temp
    copy.forEach((r, i) => (r.priority = i + 1))
    setRules(copy)
  }

  const handleDelete = () => {
    if (!selectedId) return
    const next = rules.filter((r) => r.id !== selectedId)
    next.forEach((r, i) => (r.priority = i + 1))
    setRules(next)
    setSelectedId(next[0]?.id || null)
  }

  const handleDuplicate = () => {
    if (!selectedRule) return
    const dup: CfRule = {
      ...structuredClone(selectedRule),
      id: 'cf_' + Math.random().toString(36).slice(2, 9),
      priority: rules.length + 1
    }
    const next = [...rules, dup]
    setRules(next)
    setSelectedId(dup.id)
  }

  const handleAppliesToChange = (ruleId: string, text: string) => {
    // Split by comma
    const parts = text.split(',')
    const parsedRanges = parts
      .map((p) => parseA1(p.trim()))
      .filter((p): p is NonNullable<typeof p> => p !== null)

    if (parsedRanges.length > 0) {
      setRules((prev) =>
        prev.map((r) => (r.id === ruleId ? { ...r, ranges: parsedRanges } : r))
      )
    }
  }

  const handleStopIfTrueChange = (ruleId: string, checked: boolean) => {
    setRules((prev) =>
      prev.map((r) => (r.id === ruleId ? { ...r, stopIfTrue: checked } : r))
    )
  }

  const saveToLayout = () => {
    if (store.refuseIfProtected(sheet)) return
    store.updateLayout(sheet, 'Manage Conditional Formatting', (l) => {
      l.conditionalFormats = rules
    })
  }

  const handleOk = () => {
    saveToLayout()
    onClose()
  }

  const handleApply = () => {
    saveToLayout()
  }

  return (
    <Dialog title="Conditional Formatting Rules Manager" width={680} onOk={handleOk} onCancel={onClose}>
      <div className="xl-cf-manage-body">
        <div className="xl-cf-manage-top">
          <div className="xl-cf-manage-scope">
            <span>Show formatting rules for:</span>
            <select
              className="xl-cf-select"
              value={scope}
              onChange={(e) => setScope(e.target.value as any)}
            >
              <option value="selection">Current Selection</option>
              <option value="sheet">This Worksheet</option>
            </select>
          </div>

          <div className="xl-cf-manage-actions">
            <button className="xl-cf-btn" onClick={onNewRule}>
              New Rule...
            </button>
            <button
              className="xl-cf-btn"
              disabled={!selectedRule}
              onClick={() => selectedRule && onEditRule(selectedRule)}
            >
              Edit Rule...
            </button>
            <button className="xl-cf-btn" disabled={!selectedRule} onClick={handleDelete}>
              Delete Rule
            </button>
            <button className="xl-cf-btn" disabled={!selectedRule} onClick={handleDuplicate}>
              Duplicate Rule
            </button>
            <div className="xl-cf-priority-btns">
              <button
                className="xl-cf-btn xl-cf-icon-btn"
                title="Move Up"
                disabled={!selectedRule}
                onClick={handleMoveUp}
              >
                ▲
              </button>
              <button
                className="xl-cf-btn xl-cf-icon-btn"
                title="Move Down"
                disabled={!selectedRule}
                onClick={handleMoveDown}
              >
                ▼
              </button>
            </div>
          </div>
        </div>

        <div className="xl-cf-table-container">
          <table className="xl-cf-table">
            <thead>
              <tr>
                <th style={{ width: '38%' }}>Rule (applied in order shown)</th>
                <th style={{ width: '22%' }}>Format</th>
                <th style={{ width: '26%' }}>Applies to</th>
                <th style={{ width: '14%' }}>Stop If True</th>
              </tr>
            </thead>
            <tbody>
              {visibleRules.length === 0 ? (
                <tr>
                  <td colSpan={4} className="xl-cf-empty-msg">
                    No conditional formatting rules found for this {scope}.
                  </td>
                </tr>
              ) : (
                visibleRules.map((rule) => {
                  const isSel = rule.id === selectedId
                  const fmt = rule.format
                  return (
                    <tr
                      key={rule.id}
                      className={isSel ? 'selected' : ''}
                      onClick={() => setSelectedId(rule.id)}
                      onDoubleClick={() => onEditRule(rule)}
                    >
                      <td className="xl-cf-cell-desc">{ruleDescription(rule)}</td>
                      <td className="xl-cf-cell-format">
                        {rule.type === 'colorScale' ? (
                          <div
                            className="xl-cf-preview-scale"
                            style={{
                              background: rule.colorScale?.midColor
                                ? `linear-gradient(to right, ${rule.colorScale.minColor}, ${rule.colorScale.midColor}, ${rule.colorScale.maxColor})`
                                : `linear-gradient(to right, ${rule.colorScale?.minColor || '#fff'}, ${rule.colorScale?.maxColor || '#63BE7B'})`
                            }}
                          />
                        ) : rule.type === 'dataBar' ? (
                          <div
                            className="xl-cf-preview-bar"
                            style={{
                              background: rule.dataBar?.gradient
                                ? `linear-gradient(to right, ${rule.dataBar.color}, #fff)`
                                : rule.dataBar?.color
                            }}
                          />
                        ) : rule.type === 'iconSet' ? (
                          <div className="xl-cf-preview-icons">● ● ●</div>
                        ) : (
                          <div
                            className="xl-cf-preview-chip"
                            style={{
                              backgroundColor: fmt?.fill || 'transparent',
                              color: fmt?.color || '#000',
                              fontWeight: fmt?.bold ? 'bold' : 'normal',
                              fontStyle: fmt?.italic ? 'italic' : 'normal',
                              border: fmt?.borders ? '1px solid #000' : '1px solid #ddd'
                            }}
                          >
                            AaBbCcYyZz
                          </div>
                        )}
                      </td>
                      <td className="xl-cf-cell-ranges">
                        <input
                          className="xl-cf-range-input"
                          defaultValue={rangesToString(rule)}
                          onBlur={(e) => handleAppliesToChange(rule.id, e.target.value)}
                        />
                      </td>
                      <td className="xl-cf-cell-stop">
                        <input
                          type="checkbox"
                          checked={Boolean(rule.stopIfTrue)}
                          onChange={(e) => handleStopIfTrueChange(rule.id, e.target.checked)}
                        />
                      </td>
                    </tr>
                  )
                })
              )}
            </tbody>
          </table>
        </div>

        <div className="xl-cf-manage-bottom">
          <button className="xl-cf-btn" onClick={handleApply}>
            Apply
          </button>
        </div>
      </div>
    </Dialog>
  )
}

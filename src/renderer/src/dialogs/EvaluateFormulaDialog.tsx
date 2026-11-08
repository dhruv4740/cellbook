// Formulas > Evaluate Formula: click Evaluate to replace the underlined part with its value, step by step.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { colName } from '../model/types'
import { applyStep, nextStep } from '../data/evaluateFormula'

export function EvaluateFormulaDialog({ onClose }: { onClose: () => void }) {
  const sheet = store.active
  const a = sheet.selection.active
  const raw = store.getRaw(sheet, a.r, a.c)
  const start = raw.startsWith('=') ? raw.slice(1) : ''
  const [expr, setExpr] = useState(start)
  const step = expr ? nextStep(expr, sheet) : null
  const finished = !!expr && !step
  const btn = (label: string, on: () => void, disabled = false): React.ReactElement => (
    <button className="xl-btn" style={{ display: 'block', width: '100%', marginBottom: 6 }} disabled={disabled} onClick={on}>{label}</button>
  )
  return (
    <Dialog title="Evaluate Formula" width={470} onCancel={onClose} cancelLabel="Close">
      <div style={{ display: 'flex', gap: 12, fontSize: 12 }}>
        <div style={{ flex: 1 }}>
          <div>Reference:</div>
          <div style={{ border: '1px solid #c6c6c6', padding: '3px 6px', marginBottom: 8, background: '#f7f7f7' }}>{start ? `${sheet.name}!${colName(a.c)}${a.r + 1}` : ''}</div>
          <div>Evaluation:</div>
          <div style={{ border: '1px solid #c6c6c6', minHeight: 70, padding: 6, background: '#fff', fontFamily: 'Consolas, monospace', wordBreak: 'break-all' }}>
            {!start ? <span style={{ color: '#666', fontFamily: 'inherit' }}>The active cell does not contain a formula.</span> : (
              <>
                =
                {step ? <>{expr.slice(0, step.start)}<u>{expr.slice(step.start, step.end)}</u>{expr.slice(step.end)}</> : expr}
              </>
            )}
          </div>
          {finished && <div style={{ marginTop: 6, color: '#666' }}>The cell is currently being evaluated. The result of the last evaluation is shown in italics.</div>}
        </div>
        <div style={{ width: 110, paddingTop: 20 }}>
          {finished ? btn('Restart', () => setExpr(start)) : btn('Evaluate', () => step && setExpr(applyStep(expr, step)), !step)}
          {btn('Step In', () => undefined, true)}
          {btn('Step Out', () => undefined, true)}
        </div>
      </div>
    </Dialog>
  )
}

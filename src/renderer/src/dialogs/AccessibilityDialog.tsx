// Review > Check Accessibility: a short list of the things Excel's checker flags that Cellbook can detect.
import { Dialog } from '../components/Dialog'
import { store } from '../model/store'
import { colName } from '../model/types'

interface Issue { sheet: string; where: string; what: string }

function findIssues(): Issue[] {
  const out: Issue[] = []
  for (const s of store.sheets) {
    if (s.layout.sheetHidden) continue
    if (/^Sheet\d+$/i.test(s.name)) out.push({ sheet: s.name, where: 'Sheet', what: 'Default sheet name — give the sheet a meaningful name' })
    for (const p of s.layout.pictures ?? []) if (!p.alt) out.push({ sheet: s.name, where: p.name, what: 'Missing alt text' })
    for (const sh of s.layout.shapes ?? []) if (!sh.text && sh.kind !== 'line' && sh.kind !== 'arrow') out.push({ sheet: s.name, where: sh.name, what: 'Missing alt text' })
    for (const ch of s.layout.charts ?? []) if (!ch.title) out.push({ sheet: s.name, where: `Chart at ${colName(ch.anchor.c)}${ch.anchor.r + 1}`, what: 'Chart has no title' })
    for (const m of s.layout.merges ?? []) out.push({ sheet: s.name, where: `${colName(m.c1)}${m.r1 + 1}`, what: 'Merged cell — use Center Across Selection instead' })
  }
  return out.slice(0, 200)
}

export function AccessibilityDialog({ onClose }: { onClose: () => void }) {
  const issues = findIssues()
  return (
    <Dialog title="Accessibility Checker" width={480} onCancel={onClose} cancelLabel="Close">
      <div style={{ fontSize: 12 }}>
        {!issues.length ? <div>No accessibility issues found. People with disabilities should not have difficulty reading this document.</div> : (
          <>
            <div style={{ marginBottom: 6 }}>Inspection Results ({issues.length}):</div>
            <div style={{ border: '1px solid #c6c6c6', maxHeight: 220, overflow: 'auto' }}>
              {issues.map((i, k) => <div key={k} style={{ padding: '3px 6px', borderBottom: '1px solid #eee' }}><b>{i.sheet}</b> · {i.where}: {i.what}</div>)}
            </div>
          </>
        )}
      </div>
    </Dialog>
  )
}

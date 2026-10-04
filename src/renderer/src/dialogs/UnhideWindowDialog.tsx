// View > Unhide: lists the windows hidden with View > Hide and brings the chosen one back.
import { useEffect, useState } from 'react'
import { Dialog } from '../components/Dialog'
import { MessageDialog } from './SheetDialogs'

export function UnhideWindowDialog({ onClose }: { onClose: () => void }) {
  const [list, setList] = useState<{ id: number; title: string }[] | null>(null)
  const [sel, setSel] = useState<number | null>(null)
  useEffect(() => {
    void (window.api?.listHiddenWindows?.() ?? Promise.resolve([])).then((l) => { setList(l); setSel(l[0]?.id ?? null) })
  }, [])
  if (!list) return null
  if (!list.length) return <MessageDialog text="No windows are hidden." onClose={onClose} />
  const ok = (): void => { if (sel !== null) void window.api?.activateWindow(sel); onClose() }
  return (
    <Dialog title="Unhide" width={320} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12, marginBottom: 4 }}>Unhide workbook:</div>
      <select size={8} style={{ width: '100%' }} value={sel ?? ''} onChange={(e) => setSel(+e.target.value)} onDoubleClick={ok}>
        {list.map((w) => <option key={w.id} value={w.id}>{w.title}</option>)}
      </select>
    </Dialog>
  )
}

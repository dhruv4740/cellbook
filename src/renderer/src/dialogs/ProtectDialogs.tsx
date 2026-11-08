// Review > Protect Sheet / Unprotect Sheet / Protect Workbook (Excel's dialogs and wording).
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { store, SheetState } from '../model/store'
import { setUi } from '../chrome/ui'
import { PROTECT_OPTIONS, checkPassword, defaultAllow, hashPassword } from '../model/protection'
import { SheetProtection } from '../model/types'
import { MessageDialog } from './SheetDialogs'

const WRONG_PASSWORD = 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.'

/** Asks for the password before unprotecting; `verify` checks it, `onOk` runs when it is right. */
export function PasswordPromptDialog({ title, verify, onOk, onClose }: { title: string; verify: (pw: string) => boolean; onOk: () => void; onClose: () => void }) {
  const [pw, setPw] = useState('')
  const [wrong, setWrong] = useState(false)
  if (wrong) return <MessageDialog text={WRONG_PASSWORD} onClose={() => { setWrong(false); setPw('') }} />
  return (
    <Dialog title={title} width={300} onCancel={onClose} onOk={() => { if (verify(pw)) onOk(); else setWrong(true) }}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 4 }}>Password:</div>
        <input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} style={{ width: '100%' }} />
      </div>
    </Dialog>
  )
}

export function ProtectSheetDialog({ sheetUid, onClose }: { sheetUid: number; onClose: () => void }) {
  const sheet = store.sheetByUid(sheetUid) as SheetState
  const [protect, setProtect] = useState(true)
  const [pw, setPw] = useState('')
  const [confirmPw, setConfirmPw] = useState<string | null>(null)
  const [allow, setAllow] = useState(defaultAllow())
  const [mismatch, setMismatch] = useState(false)

  const apply = (password: string) => {
    const prot: SheetProtection = { allow, ...(password ? { password: hashPassword(password) } : {}) }
    store.setSheetProtection(sheet, prot)
    onClose()
  }
  if (mismatch) return <MessageDialog text="Confirmation password is not identical." onClose={() => setMismatch(false)} />
  if (confirmPw !== null) {
    return (
      <Dialog title="Confirm Password" width={330} onCancel={onClose} onOk={() => { if (confirmPw === pw) apply(pw); else { setMismatch(true); setConfirmPw('') } }}>
        <div style={{ fontSize: 12 }}>
          <div style={{ marginBottom: 4 }}>Reenter password to proceed.</div>
          <input type="password" autoFocus value={confirmPw} onChange={(e) => setConfirmPw(e.target.value)} style={{ width: '100%' }} />
          <div style={{ marginTop: 8 }}>Caution: If you lose or forget the password, it cannot be recovered. It is advisable to keep a list of passwords and their corresponding workbook and sheet names in a safe place. (Remember that passwords are case-sensitive.)</div>
        </div>
      </Dialog>
    )
  }
  const ok = () => {
    if (!protect) { onClose(); return }
    if (pw) setConfirmPw('')
    else apply('')
  }
  return (
    <Dialog title="Protect Sheet" width={340} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <label style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
          <input type="checkbox" checked={protect} onChange={(e) => setProtect(e.target.checked)} /> Protect worksheet and contents of locked cells
        </label>
        <div style={{ margin: '8px 0 4px' }}>Password to unprotect sheet:</div>
        <input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} style={{ width: '100%' }} disabled={!protect} />
        <div style={{ margin: '8px 0 4px' }}>Allow all users of this worksheet to:</div>
        <div style={{ border: '1px solid #8a8a8a', height: 190, overflowY: 'auto', background: '#fff', padding: 2 }}>
          {PROTECT_OPTIONS.map((o) => (
            <label key={o.key} style={{ display: 'flex', gap: 6, alignItems: 'center', padding: '1px 2px' }}>
              <input type="checkbox" disabled={!protect} checked={allow[o.key]} onChange={(e) => setAllow({ ...allow, [o.key]: e.target.checked })} /> {o.label}
            </label>
          ))}
        </div>
      </div>
    </Dialog>
  )
}

/** Unprotect Sheet: asks for the password when there is one, otherwise unprotects straight away. */
export function unprotectSheet(sheet: SheetState): void {
  const p = sheet.layout.protection
  if (!p) return
  if (!p.password) { store.setSheetProtection(sheet, null); return }
  setUi({ dialog: { kind: 'password', props: { title: 'Unprotect Sheet', verify: (pw: string) => checkPassword(p.password, pw), onOk: () => { store.setSheetProtection(sheet, null); setUi({ dialog: null }) } } } })
}

export function ProtectWorkbookDialog({ onClose }: { onClose: () => void }) {
  const [structure, setStructure] = useState(true)
  const [windows, setWindows] = useState(false)
  const [pw, setPw] = useState('')
  const [confirmPw, setConfirmPw] = useState<string | null>(null)
  const [mismatch, setMismatch] = useState(false)
  const apply = (password: string) => {
    store.setBookProtection({ lockStructure: structure, lockWindows: windows, ...(password ? { password: hashPassword(password) } : {}) })
    onClose()
  }
  if (mismatch) return <MessageDialog text="Confirmation password is not identical." onClose={() => setMismatch(false)} />
  if (confirmPw !== null) {
    return (
      <Dialog title="Confirm Password" width={330} onCancel={onClose} onOk={() => { if (confirmPw === pw) apply(pw); else { setMismatch(true); setConfirmPw('') } }}>
        <div style={{ fontSize: 12 }}>
          <div style={{ marginBottom: 4 }}>Reenter password to proceed.</div>
          <input type="password" autoFocus value={confirmPw} onChange={(e) => setConfirmPw(e.target.value)} style={{ width: '100%' }} />
          <div style={{ marginTop: 8 }}>Caution: If you lose or forget the password, it cannot be recovered. It is advisable to keep a list of passwords and their corresponding workbook and sheet names in a safe place. (Remember that passwords are case-sensitive.)</div>
        </div>
      </Dialog>
    )
  }
  return (
    <Dialog title="Protect Structure and Windows" width={320} onCancel={onClose} onOk={() => { if (!structure && !windows) { onClose(); return } if (pw) setConfirmPw(''); else apply('') }}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 4 }}>Password (optional):</div>
        <input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} style={{ width: '100%' }} />
        <div style={{ margin: '8px 0 2px' }}>Protect workbook for</div>
        <label style={{ display: 'flex', gap: 6, alignItems: 'center' }}><input type="checkbox" checked={structure} onChange={(e) => setStructure(e.target.checked)} /> Structure</label>
        <label style={{ display: 'flex', gap: 6, alignItems: 'center' }}><input type="checkbox" checked={windows} onChange={(e) => setWindows(e.target.checked)} /> Windows</label>
      </div>
    </Dialog>
  )
}

export function unprotectWorkbook(): void {
  const p = store.bookProtection
  if (!p) return
  if (!p.password) { store.setBookProtection(undefined); return }
  setUi({ dialog: { kind: 'password', props: { title: 'Unprotect Workbook', verify: (pw: string) => checkPassword(p.password, pw), onOk: () => { store.setBookProtection(undefined); setUi({ dialog: null }) } } } })
}

export function toggleSheetProtection(sheet: SheetState): void {
  if (sheet.layout.protection) unprotectSheet(sheet)
  else setUi({ dialog: { kind: 'protectSheet', props: { sheetUid: sheet.uid } } })
}
export function toggleWorkbookProtection(): void {
  if (store.bookProtection) unprotectWorkbook()
  else setUi({ dialog: { kind: 'protectWorkbook' } })
}

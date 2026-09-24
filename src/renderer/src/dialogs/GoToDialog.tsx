import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { goTo } from '../grid/commands'
import './GoToDialog.css'

export interface GoToDialogProps {
  onOpenSpecial: () => void
  onClose: () => void
}

// Module-level history so it persists across dialog openings
const recentRefs: string[] = ['$A$1']

export function GoToDialog({ onOpenSpecial, onClose }: GoToDialogProps) {
  const [refText, setRefText] = useState('')
  const [selectedItem, setSelectedItem] = useState<string | null>(null)
  const [errorMsg, setErrorMsg] = useState<string | null>(null)

  const handleOk = () => {
    const target = (refText || selectedItem || '').trim()
    if (!target) return
    const success = goTo(target)
    if (success) {
      if (!recentRefs.includes(target)) {
        recentRefs.unshift(target)
        if (recentRefs.length > 20) recentRefs.pop()
      }
      onClose()
    } else {
      setErrorMsg("Reference is not valid.")
    }
  }

  return (
    <Dialog
      title="Go To"
      width={320}
      onOk={handleOk}
      onCancel={onClose}
      footerLeft={
        <button
          className="xl-btn"
          style={{ minWidth: 70 }}
          onClick={() => {
            onClose()
            onOpenSpecial()
          }}
        >
          <u>S</u>pecial...
        </button>
      }
    >
      <div className="xl-goto-body">
        <div style={{ fontWeight: 500 }}><u>G</u>o to:</div>
        <div className="xl-goto-listbox">
          {recentRefs.map((item) => (
            <div
              key={item}
              className={`xl-goto-item ${selectedItem === item ? 'selected' : ''}`}
              onClick={() => {
                setSelectedItem(item)
                setRefText(item)
              }}
              onDoubleClick={() => {
                setRefText(item)
                goTo(item)
                onClose()
              }}
            >
              {item}
            </div>
          ))}
        </div>

        <div style={{ fontWeight: 500, marginTop: 4 }}><u>R</u>eference:</div>
        <input
          className="xl-goto-input"
          value={refText}
          autoFocus
          onChange={(e) => {
            setRefText(e.target.value)
            setSelectedItem(null)
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              handleOk()
            }
          }}
        />

        {errorMsg && (
          <div style={{ color: '#d83b01', fontSize: 11, marginTop: 2 }}>
            {errorMsg}
          </div>
        )}
      </div>
    </Dialog>
  )
}

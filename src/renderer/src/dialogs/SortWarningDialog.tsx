import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import './SortDialog.css'

export interface SortWarningDialogProps {
  onSort: (expand: boolean) => void
  onClose: () => void
}

export function SortWarningDialog({ onSort, onClose }: SortWarningDialogProps) {
  const [choice, setChoice] = useState<'expand' | 'current'>('expand')

  return (
    <Dialog
      title="Sort Warning"
      width={400}
      okLabel="Sort"
      onOk={() => {
        onSort(choice === 'expand')
        onClose()
      }}
      onCancel={onClose}
    >
      <div className="xl-sort-warning-body">
        <p className="xl-sort-warning-text">
          Microsoft Excel found data next to your selection. Since you have not selected this data, it will not be sorted.
        </p>
        <p className="xl-sort-warning-prompt">What do you want to do?</p>
        <div className="xl-sort-warning-options">
          <label className="xl-sort-radio-label">
            <input
              type="radio"
              name="sortOption"
              checked={choice === 'expand'}
              onChange={() => setChoice('expand')}
            />
            <span>Expand the selection</span>
          </label>
          <label className="xl-sort-radio-label">
            <input
              type="radio"
              name="sortOption"
              checked={choice === 'current'}
              onChange={() => setChoice('current')}
            />
            <span>Continue with the current selection</span>
          </label>
        </div>
      </div>
    </Dialog>
  )
}

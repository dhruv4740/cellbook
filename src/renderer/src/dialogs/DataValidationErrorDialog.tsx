import { Dialog } from '../components/Dialog'
import { DvErrorStyle } from '../model/types'
import { ErrorCircle20Regular, Warning20Regular, Info20Regular } from '@fluentui/react-icons'

export interface DataValidationErrorDialogProps {
  style?: DvErrorStyle
  title?: string
  message?: string
  onRetry: () => void
  onCancel: () => void
  onContinue: () => void
}

export function DataValidationErrorDialog({
  style = 'stop',
  title = 'Microsoft Excel',
  message = "This value doesn't match the data validation restrictions defined for this cell.",
  onRetry,
  onCancel,
  onContinue
}: DataValidationErrorDialogProps) {
  const icon =
    style === 'stop' ? (
      <ErrorCircle20Regular style={{ color: '#d13438', fontSize: 32 }} />
    ) : style === 'warning' ? (
      <Warning20Regular style={{ color: '#ffaa44', fontSize: 32 }} />
    ) : (
      <Info20Regular style={{ color: '#0078d7', fontSize: 32 }} />
    )

  return (
    <Dialog title={title} width={380} hideFooter onCancel={onCancel}>
      <div style={{ display: 'flex', gap: 16, alignItems: 'flex-start', padding: '10px 4px' }}>
        <div>{icon}</div>
        <div style={{ flex: 1, font: '12px var(--ui-font)', lineHeight: 1.4 }}>{message}</div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
        {style === 'stop' && (
          <>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              autoFocus
              onClick={onRetry}
            >
              Retry
            </button>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              onClick={onCancel}
            >
              Cancel
            </button>
          </>
        )}
        {style === 'warning' && (
          <>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              onClick={onContinue}
            >
              Yes
            </button>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              autoFocus
              onClick={onRetry}
            >
              No
            </button>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              onClick={onCancel}
            >
              Cancel
            </button>
          </>
        )}
        {style === 'information' && (
          <>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              autoFocus
              onClick={onContinue}
            >
              OK
            </button>
            <button
              style={{ minWidth: 68, height: 24, padding: '0 10px', font: '12px var(--ui-font)' }}
              onClick={onCancel}
            >
              Cancel
            </button>
          </>
        )}
      </div>
    </Dialog>
  )
}

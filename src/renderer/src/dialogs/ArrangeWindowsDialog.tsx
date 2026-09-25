import { useState } from 'react'
import { Dialog } from '../components/Dialog'

export type ArrangeMode = 'tiled' | 'horizontal' | 'vertical' | 'cascade'

export function ArrangeWindowsDialog({ onClose }: { onClose: () => void }) {
  const [mode, setMode] = useState<ArrangeMode>('tiled')
  const [activeWorkbookOnly, setActiveWorkbookOnly] = useState(false)

  return (
    <Dialog
      title="Arrange Windows"
      width={320}
      okLabel="OK"
      onOk={() => {
        void window.api?.arrangeWindows(mode)
        onClose()
      }}
      onCancel={onClose}
    >
      <div style={{ padding: '8px 4px', display: 'flex', flexDirection: 'column', gap: '12px', fontSize: '12px', color: '#242424' }}>
        <fieldset style={{ border: '1px solid #d1d1d1', padding: '10px 14px', borderRadius: '2px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
          <legend style={{ padding: '0 4px', color: '#333' }}>Arrange</legend>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'default' }}>
            <input type="radio" name="arrangeMode" checked={mode === 'tiled'} onChange={() => setMode('tiled')} />
            <span>Tiled</span>
          </label>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'default' }}>
            <input type="radio" name="arrangeMode" checked={mode === 'horizontal'} onChange={() => setMode('horizontal')} />
            <span>Horizontal</span>
          </label>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'default' }}>
            <input type="radio" name="arrangeMode" checked={mode === 'vertical'} onChange={() => setMode('vertical')} />
            <span>Vertical</span>
          </label>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'default' }}>
            <input type="radio" name="arrangeMode" checked={mode === 'cascade'} onChange={() => setMode('cascade')} />
            <span>Cascade</span>
          </label>
        </fieldset>
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'default' }}>
          <input type="checkbox" checked={activeWorkbookOnly} onChange={(e) => setActiveWorkbookOnly(e.target.checked)} />
          <span>Windows of active workbook</span>
        </label>
      </div>
    </Dialog>
  )
}

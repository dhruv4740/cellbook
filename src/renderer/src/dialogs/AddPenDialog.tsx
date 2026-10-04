// Draw > Add > Add Pen: a named pen with a colour and thickness; it joins the Drawing Tools row for this session.
import { useState } from 'react'
import { Dialog } from '../components/Dialog'
import { getUi, setUi } from '../chrome/ui'
import { Pen } from '../data/ink'

export function AddPenDialog({ onClose }: { onClose: () => void }) {
  const [color, setColor] = useState('#7030a0')
  const [width, setWidth] = useState(3)
  const [highlighter, setHighlighter] = useState(false)
  const ok = (): void => {
    const pen: Pen = { name: `Pen ${getUi().customPens.length + 1}`, color, width: highlighter ? Math.max(width, 8) : width, ...(highlighter ? { alpha: 0.5 } : {}) }
    setUi({ customPens: [...getUi().customPens, pen], inkTool: { kind: 'pen', pen }, selectObjectsMode: false })
    onClose()
  }
  return (
    <Dialog title="Add Pen" width={320} onCancel={onClose} onOk={ok}>
      <div style={{ fontSize: 12 }}>
        <div className="xl-field-row"><label style={{ width: 90 }}>Color:</label><input type="color" value={color} onChange={(e) => setColor(e.target.value)} /></div>
        <div className="xl-field-row" style={{ marginTop: 6 }}>
          <label style={{ width: 90 }}>Thickness:</label>
          <input type="range" min={0.5} max={20} step={0.5} value={width} onChange={(e) => setWidth(Number(e.target.value))} />
          <span style={{ marginLeft: 6, width: 40 }}>{width} px</span>
        </div>
        <label style={{ display: 'block', marginTop: 6 }}><input type="checkbox" checked={highlighter} onChange={(e) => setHighlighter(e.target.checked)} /> Highlighter (semi-transparent)</label>
        <div style={{ marginTop: 10, height: 30, border: '1px solid #d0d0d0', display: 'flex', alignItems: 'center', padding: '0 8px' }}>
          <div style={{ width: '100%', height: highlighter ? Math.max(width, 8) : width, background: color, opacity: highlighter ? 0.5 : 1, borderRadius: 4 }} />
        </div>
      </div>
    </Dialog>
  )
}

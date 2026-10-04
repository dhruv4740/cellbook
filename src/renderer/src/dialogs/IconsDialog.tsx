// Insert > Icons: searchable, categorized icon picker; the chosen icon goes in as a picture.
import { useRef, useState } from 'react'
import { Dialog } from '../components/Dialog'
import { ICON_CATEGORIES } from '../data/iconSet'
import { searchIcons, insertIcon } from '../data/icons'

export function IconsDialog({ onClose }: { onClose: () => void }) {
  const [query, setQuery] = useState('')
  const [cat, setCat] = useState<string | null>(null)
  const [pick, setPick] = useState<string | null>(null)
  const [color, setColor] = useState('#000000')
  const grid = useRef<HTMLDivElement>(null)
  const found = searchIcons(query, cat)
  const insert = (name: string): void => {
    const svg = Array.from(grid.current?.querySelectorAll<HTMLElement>('[data-icon]') ?? []).find((el) => el.dataset.icon === name)?.querySelector('svg')
    if (!svg) return
    void insertIcon(svg.outerHTML, color, name).then(onClose)
  }
  const ok = (): void => { if (pick) insert(pick) }
  return (
    <Dialog title="Insert Icons" width={560} onCancel={onClose} onOk={ok} okLabel="Insert">
      <div style={{ fontSize: 12 }}>
        <div className="xl-field-row" style={{ marginBottom: 6, gap: 6 }}>
          <input autoFocus placeholder="Search icons" style={{ flex: 1 }} value={query} onChange={(e) => setQuery(e.target.value)} />
          <label>Color: <input type="color" value={color} onChange={(e) => setColor(e.target.value)} /></label>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <div style={{ width: 120, display: 'flex', flexDirection: 'column' }}>
            {[null, ...ICON_CATEGORIES.map((c) => c.name)].map((c) => (
              <div key={c ?? 'all'} onClick={() => setCat(c)} style={{ padding: '3px 6px', cursor: 'default', background: cat === c ? '#e2f0e8' : undefined, fontWeight: cat === c ? 600 : 400 }}>{c ?? 'All'}</div>
            ))}
          </div>
          <div ref={grid} style={{ display: 'flex', flexWrap: 'wrap', alignContent: 'flex-start', height: 280, flex: 1, overflow: 'auto' }}>
            {found.map(({ name, Icon }) => (
              <div key={name} data-icon={name} title={name} onClick={() => setPick(name)} onDoubleClick={() => insert(name)}
                style={{ width: 48, height: 48, display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid ' + (pick === name ? '#217346' : 'transparent'), background: pick === name ? '#e2f0e8' : undefined }}>
                <Icon style={{ width: 28, height: 28, color }} />
              </div>
            ))}
            {!found.length && <div style={{ color: '#666', padding: 8 }}>No icons match.</div>}
          </div>
        </div>
      </div>
    </Dialog>
  )
}

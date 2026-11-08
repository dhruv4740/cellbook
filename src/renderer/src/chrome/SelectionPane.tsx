// Page Layout > Selection Pane: lists the active sheet's floating objects front to back, click to select,
// eye to show / hide, arrows to change the stacking order.
import { store, useStore } from '../model/store'
import { setUi, useUi } from './ui'
import { floatsOf, reorderIds } from '../data/arrange'
import { selectedObjectIds, selectObject, setObjectsHidden, setObjectOrder } from '../data/arrangeActions'

export default function SelectionPane() {
  useStore()
  const ui = useUi()
  const sheet = store.active
  const items = floatsOf(sheet.layout).reverse() // front first, like Excel
  const sel = selectedObjectIds()
  const btn = { cursor: 'pointer', padding: '0 4px', userSelect: 'none' } as const
  const move = (id: string, mode: 'forward' | 'backward'): void => setObjectOrder(reorderIds(floatsOf(sheet.layout).map((f) => f.id), [id], mode))
  const all = items.map((f) => f.id)
  return (
    <div style={{ position: 'fixed', right: 12, top: 190, width: 250, background: '#fff', border: '1px solid #8a8a8a', boxShadow: '0 4px 16px rgba(0,0,0,0.25)', zIndex: 900, fontSize: 12 }}>
      <div style={{ display: 'flex', alignItems: 'center', background: '#f3f3f3', padding: '4px 8px', userSelect: 'none' }}>
        <span style={{ flex: 1 }}>Selection</span>
        <span role="button" aria-label="Close Selection Pane" style={btn} onClick={() => setUi({ selectionPaneOpen: false })}>✕</span>
      </div>
      <div style={{ padding: 6, display: 'flex', gap: 6 }}>
        <button className="xl-btn" disabled={!items.length} onClick={() => setObjectsHidden(all, false)}>Show All</button>
        <button className="xl-btn" disabled={!items.length} onClick={() => setObjectsHidden(all, true)}>Hide All</button>
      </div>
      <div style={{ maxHeight: 300, overflowY: 'auto', borderTop: '1px solid #d0d0d0' }}>
        {!items.length && <div style={{ padding: 8, color: '#666' }}>No objects on this sheet.</div>}
        {items.map((f) => (
          <div key={f.id} onClick={() => !f.obj.hidden && selectObject(sheet, f.id)}
            style={{ display: 'flex', alignItems: 'center', padding: '3px 6px', background: sel.includes(f.id) ? '#dcebe3' : undefined, opacity: f.obj.hidden ? 0.5 : 1, cursor: 'default' }}>
            <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}{f.obj.group ? ' (grouped)' : ''}</span>
            <span role="button" aria-label={f.obj.hidden ? 'Show' : 'Hide'} title={f.obj.hidden ? 'Show' : 'Hide'} style={btn} onClick={(e) => { e.stopPropagation(); setObjectsHidden([f.id], !f.obj.hidden) }}>{f.obj.hidden ? '—' : '◉'}</span>
            <span role="button" aria-label="Bring Forward" title="Bring Forward" style={btn} onClick={(e) => { e.stopPropagation(); move(f.id, 'forward') }}>▲</span>
            <span role="button" aria-label="Send Backward" title="Send Backward" style={btn} onClick={(e) => { e.stopPropagation(); move(f.id, 'backward') }}>▼</span>
          </div>
        ))}
      </div>
    </div>
  )
}

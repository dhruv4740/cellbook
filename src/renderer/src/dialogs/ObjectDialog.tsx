// Insert > Object. OLE embedding needs Office on Windows, so Create from File embeds image files as pictures.
import { Dialog } from '../components/Dialog'
import { pickPictureFiles } from '../data/pictures'

export function ObjectDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog title="Object" width={420} onCancel={onClose} onOk={undefined}
      footerLeft={<button className="xl-btn" data-object-file onClick={() => { onClose(); pickPictureFiles() }}>Create from File…</button>}>
      <div style={{ fontSize: 12 }}>
        <div style={{ marginBottom: 6 }}>Create New</div>
        <div style={{ border: '1px solid #d1d1d1', padding: 8, color: '#605e5c' }}>Cellbook can’t create linked Office objects (OLE). Use Create from File to embed an image in the sheet.</div>
      </div>
    </Dialog>
  )
}

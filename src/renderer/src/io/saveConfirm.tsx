// Imperative "Want to save your changes?" prompt, mountable without any change to App.tsx.
import { createRoot } from 'react-dom/client'
import { Dialog } from '../components/Dialog'

export type SaveChoice = 'save' | 'dontsave' | 'cancel'

/** Shows Excel's save-changes prompt and resolves with the user's choice. */
export function confirmSaveChanges(fileName: string): Promise<SaveChoice> {
  return new Promise((resolve) => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const close = (choice: SaveChoice): void => {
      root.unmount()
      host.remove()
      resolve(choice)
    }
    root.render(
      <Dialog
        title="Microsoft Excel"
        width={380}
        onOk={() => close('save')}
        okLabel="Save"
        onCancel={() => close('cancel')}
        cancelLabel="Cancel"
        footerLeft={
          <button className="xl-btn" onClick={() => close('dontsave')}>
            Don&apos;t Save
          </button>
        }
      >
        <div style={{ padding: '4px 4px 12px', fontSize: 13, lineHeight: 1.5 }}>
          Want to save your changes to &apos;{fileName}&apos;?
        </div>
      </Dialog>
    )
  })
}

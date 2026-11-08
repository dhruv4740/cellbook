// Sheet commands shared by the sheet-tab menu and Home > Format.
import { store } from '../model/store'
import { setUi } from './ui'

export function hideSheet(index = store.activeIndex): void {
  const err = store.hideSheet(index)
  if (err) setUi({ dialog: { kind: 'message', props: { text: err } } })
}

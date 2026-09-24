/// <reference types="vite/client" />
import type { Api } from '../../preload'

declare global {
  interface Window {
    /** Electron bridge; undefined when running in a plain browser (npm run web). */
    api?: Api
  }
}

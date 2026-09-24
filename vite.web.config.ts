// Runs the renderer in a plain browser (no Electron) for quick UI testing.
import { resolve } from 'path'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  root: 'src/renderer',
  resolve: { alias: { '@': resolve('src/renderer/src') } },
  plugins: [react()],
  server: { port: 5173 }
})

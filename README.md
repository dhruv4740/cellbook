# Cellbook

A desktop spreadsheet app that mirrors Microsoft Excel (Microsoft 365, Windows) — same ribbon, menus, names and behavior.

## Run

```bash
npm install
npm run dev        # desktop app (Electron)
npm run web        # renderer only, in a browser at http://localhost:5173
npm run typecheck
npm test
npm run build      # production build into out/
```

## Layout

- `src/main`, `src/preload` — Electron window, file dialogs, IPC.
- `src/renderer/src/model` — workbook store (contents + formulas via HyperFormula, styles, layout, undo/redo).
- `src/renderer/src/grid` — canvas grid, selection, editing, clipboard, fill, context menus, commands.
- `src/renderer/src/engine` — formula reference shifting, autofill series.
- `src/renderer/src/chrome` — title bar, ribbon, formula bar, sheet tabs, status bar.
- `src/renderer/src/format` — Excel number-format engine.
- `src/renderer/src/io` — .xlsx/.csv read/write (exceljs), File backstage.

The formula engine is HyperFormula (GPL-3.0).

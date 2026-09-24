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

## Windows installer

`npm run dist` builds a Windows NSIS installer ("Cellbook Setup &lt;version&gt;.exe") via
electron-builder, using the `build` config in `package.json`. Building the Windows `.exe`
requires **wine** and can only be done on Windows itself (or macOS/Linux with wine installed) —
electron-builder's NSIS target shells out to `signtool.exe`/NSIS via wine, which this
Linux container doesn't have. Run `npm run dist` on a Windows machine (or a Linux/macOS
machine with wine + mono installed) to produce the real installer.

To sanity-check the electron-builder config itself without wine, `npm run dist:dir` builds an
unpacked Linux app into `release/linux-unpacked/` — it exercises packaging, native module
rebuilding and file associations config, just not the NSIS installer step.

The installer offers a per-machine or per-user install, Start Menu + desktop shortcuts, and
registers `.xlsx`, `.xlsm`, `.xls` and `.csv` file associations. It currently uses Electron's
default icon — **TODO: add a real `build/icon.ico`** (and reference it from `build.win.icon`
in `package.json`) before shipping to the pilot machines.

## Layout

- `src/main`, `src/preload` — Electron window, file dialogs, IPC.
- `src/renderer/src/model` — workbook store (contents + formulas via HyperFormula, styles, layout, undo/redo).
- `src/renderer/src/grid` — canvas grid, selection, editing, clipboard, fill, context menus, commands.
- `src/renderer/src/engine` — formula reference shifting, autofill series.
- `src/renderer/src/chrome` — title bar, ribbon, formula bar, sheet tabs, status bar.
- `src/renderer/src/format` — Excel number-format engine.
- `src/renderer/src/io` — .xlsx/.csv read/write (exceljs), File backstage.

The formula engine is HyperFormula (GPL-3.0).

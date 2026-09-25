# Cellbook status and resume notes

Target: an Excel 2021 (Windows) clone for Raman Roadways (Electron + React + TypeScript, HyperFormula engine).
Branch `milestone-3` holds everything after PR #3. Checkpoint taken on 2026-09-25 because of usage limits.

## Done and tested
- Grid, editing, undo, copy/paste, fill, merges, freeze panes, hide/unhide, zoom, sheet tabs.
- Excel 2021 look: ribbon with KeyTips on every tab (Alt, F10, legacy Alt+E+S style sequences), shortcut list in `docs/shortcuts.md`.
- One window per workbook; Ctrl+Tab/F6; spare window for instant Ctrl+N; Windows installer is built by CI (`windows-installer` job artifact).
- Fast .xlsx reader in a Web Worker (`io/fastXlsx.ts`); the legacy exceljs reader is kept for differential tests.
- Format Cells dialog, Find/Replace, Go To / Go To Special (Gemini).
- Conditional formatting: all rule types, including Solid Fill data bars (x14 extension). Saved and read back (`io/xlsxCfDv.ts`).
- Data tools: Sort dialog, Remove Duplicates, Text to Columns, Data Validation with in-cell list dropdown (Gemini). Validation is saved to .xlsx.
- AutoFilter with saved criteria; notes; threaded comments, saved as Excel's legacy fallback note.
- Charts: column/bar/line/area/pie/doughnut/scatter. Saved as real DrawingML and read from Excel files (`io/xlsxCharts.ts`). 3D, combo, radar, stock, surface and bubble charts are skipped on read.
- Macros in .xlsm are kept on save (no VBA execution).
- Excel Tables: Ctrl+T/Ctrl+L, Format as Table gallery, Table Design tab, Total Row, auto-expand, calculated columns, Convert to Range, saved to .xlsx (`model/tables.ts`, `io/xlsxTables.ts`).
- Printing: Page Layout tab, the 4-tab Page Setup dialog with custom header/footer, File > Print pane with live preview and printer list, Export to PDF, Page Break Preview overlay, page setup saved to .xlsx (`print/`, `io/xlsxPageSetup.ts`).

## Partly done: resume here
1. **Old .xls files** (`io/xls/`, uses the `cfb` package): the BIFF8 reader opens values, styles, formulas, merges, frozen panes and notes, and marks the title "[Compatibility Mode]". Saving goes to Save As .xlsx. Gaps:
   - Only about 150 of BIFF8's ~370 functions are mapped; a formula with an unmapped function keeps its cached value instead.
   - Tab colors are not read.
   - Hidden sheets show as normal tabs, because the app has no hidden-sheet concept yet.
   - Note text keeps an "Author:" prefix.
2. **Printing gaps**:
   - Screenshots and real Electron print/PDF checks were not done.
   - Page Layout view has no canvas effect yet.
   - Page breaks can't be dragged.
   - Pages from/to isn't wired.
   - Header/footer &B/&I/&U and font codes are stripped.
3. **Pictures, shapes, text boxes**: only types exist (`DrawingObject` in `model/types.ts`). **Files with logos or shapes lose them on save until this is built.** Plan, in order:
   - drawings store
   - overlay (like `charts/ChartOverlay.tsx`)
   - Insert tab wiring, plus Picture Format (JP) and Shape Format (JD) tabs
   - read/write in `io/xlsxCharts.ts` + `fastXlsx.ts`, preserving unsupported anchors verbatim
   - a print hook
4. **Functions**:
   - HyperFormula 3.4 already has XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE, IFS, SWITCH, TEXTJOIN, MAXIFS and others.
   - Missing: XMATCH, SORTBY, RANDARRAY, LET, CONCAT, NUMBERVALUE, AGGREGATE, INDIRECT, CELL, INFO, and the bond/amortization financial functions.
   - SUBTOTAL doesn't skip rows hidden by our filter; `engine/functions/hiddenRowsBridge.ts` was started for this and isn't wired in.
   - Not built: function AutoComplete and argument ScreenTips, the Insert Function / Function Arguments dialogs, and `docs/functions.md`.
   - A worker reported that HyperFormula's SORT/UNIQUE/FILTER source inside `node_modules` looks hand-patched. Verify with a clean `npm ci` before relying on it.
5. **Tables**: only one table per sheet can show filter buttons (they share the sheet AutoFilter). Table styles are approximated from theme colors.

## Not started
- Chart sheets (F11).
- View > New Window on the same workbook.
- PivotTables.
- Wiring the many ribbon buttons that do nothing yet: Insert, Review (protect sheet/workbook), View, Formulas auditing, and others.
- Full QA pass on the Windows build.

## How to work on this
- `npm run web` runs the renderer in a browser.
- Checks: `npx tsc --noEmit -p tsconfig.json` and `npx vitest run`. Running tests rewrites `samples/sample.xlsx`; restore it before committing.
- Gemini prompts go in `docs/gemini-prompts/`.

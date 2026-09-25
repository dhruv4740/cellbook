You are working on "cellbook", a desktop Excel clone (Electron + React 19 + TypeScript + HyperFormula) at https://github.com/dhruv4740/cellbook. Another AI (Claude) reviews and merges your branch, so your work must end up pushed on your own branch. Goal of the app: look and behave EXACTLY like Microsoft Excel for Microsoft 365 on Windows, and stay fast on very large files (hundreds of thousands of rows).

## Setup

```bash
git clone https://github.com/dhruv4740/cellbook   # or: cd cellbook && git fetch origin
cd cellbook
git checkout claude/optimistic-einstein-x3mu90 && git pull
git checkout -b gemini/new-window
npm install
```
Base is `claude/optimistic-einstein-x3mu90` (NOT main; it holds the latest merged work). Run with `npm run web` or `npm run dev`. Before pushing, `npm run typecheck` and `npm test` must pass and you must check the feature by hand. Don't commit `samples/sample.xlsx` (`git restore samples/sample.xlsx`). All edits go through store methods (`model/store.ts`) so they are undoable. Keep edits outside your task minimal. Put pure logic in `.ts` files with vitest tests next to them.

.xlsx files: reading uses the FAST reader `src/renderer/src/io/fastXlsx.ts` (hand-written XML scanner, runs in a worker; see how `parseConditionalFormats`/`parseDxfs` were added there). The exceljs reader in `io/xlsx.ts` (`readXlsxExceljs`) is legacy; writing uses exceljs in `writeXlsx` in `io/xlsx.ts` plus `io/xlsxPostProcess.ts` for raw XML patches exceljs can't do. Any new saved feature needs: write + fast read + a round-trip test in `io/xlsx.test.ts` (write then `readXlsx`) + it must open correctly in real Excel (valid OOXML).

When done: commit, `git push -u origin gemini/new-window`, and reply with a short summary and anything unfinished.

## Task: View > New Window (Excel behaviour)

The button exists (`chrome/ribbon/OtherTabs.tsx`, `window.api.newWindow` in `src/preload/index.ts`, handler `window:new` in `src/main/index.ts`) but currently opens a blank workbook. In Excel, View > New Window opens a SECOND view of the SAME workbook:
1. Both windows show the same workbook; an edit in one appears in the other immediately (same undo history). Each window keeps its own active sheet, selection, scroll, zoom, freeze panes view.
2. Title bars become `Book1.xlsx:1` and `Book1.xlsx:2` (reverts to plain name when one closes).
3. Saving from either saves the one workbook; closing one window doesn't prompt to save, closing the last one does.
4. Also enable View > Arrange All (Tiled / Horizontal / Vertical / Cascade dialog) and View > Switch Windows listing open windows.
Design: keep one authoritative workbook in the main process or the first renderer, and sync edits as store operations over IPC (don't copy the whole workbook per edit - files can have 500k+ rows). Add tests for the sync logic.

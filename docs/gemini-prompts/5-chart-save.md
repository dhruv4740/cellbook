You are working on "cellbook", a desktop Excel clone (Electron + React 19 + TypeScript + HyperFormula) at https://github.com/dhruv4740/cellbook. Another AI (Claude) reviews and merges your branch, so your work must end up pushed on your own branch. Goal of the app: look and behave EXACTLY like Microsoft Excel for Microsoft 365 on Windows, and stay fast on very large files (hundreds of thousands of rows).

## Setup

```bash
git clone https://github.com/dhruv4740/cellbook   # or: cd cellbook && git fetch origin
cd cellbook
git checkout claude/optimistic-einstein-x3mu90 && git pull
git checkout -b gemini/chart-save
npm install
```
Base is `claude/optimistic-einstein-x3mu90` (NOT main; it holds the latest merged work). Run with `npm run web` or `npm run dev`. Before pushing, `npm run typecheck` and `npm test` must pass and you must check the feature by hand. Don't commit `samples/sample.xlsx` (`git restore samples/sample.xlsx`). All edits go through store methods (`model/store.ts`) so they are undoable. Keep edits outside your task minimal. Put pure logic in `.ts` files with vitest tests next to them.

.xlsx files: reading uses the FAST reader `src/renderer/src/io/fastXlsx.ts` (hand-written XML scanner, runs in a worker; see how `parseConditionalFormats`/`parseDxfs` were added there). The exceljs reader in `io/xlsx.ts` (`readXlsxExceljs`) is legacy; writing uses exceljs in `writeXlsx` in `io/xlsx.ts` plus `io/xlsxPostProcess.ts` for raw XML patches exceljs can't do. Any new saved feature needs: write + fast read + a round-trip test in `io/xlsx.test.ts` (write then `readXlsx`) + it must open correctly in real Excel (valid OOXML).

When done: commit, `git push -u origin gemini/chart-save`, and reply with a short summary and anything unfinished.

## Task: Save and load charts in .xlsx

Charts exist in the app (`src/renderer/src/charts/*`, chart types in `model/types.ts`) but are not written to or read from .xlsx.
1. Write each chart as real DrawingML: `xl/drawings/drawingN.xml` (twoCellAnchor using the chart's anchor cells/offsets, `graphicFrame` → `c:chart r:id`), `xl/charts/chartN.xml` (`c:chartSpace` with the correct plot element per type: barChart/barDir col|bar, lineChart, pieChart, areaChart, scatterChart, doughnutChart, etc. as supported by the model; series `c:tx`, `c:cat`, `c:val` / `c:xVal`/`c:yVal` as `c:numRef`/`c:strRef` formulas like `Sheet1!$B$2:$B$10` with a cache; title, legend position, axes), the sheet's `<drawing r:id>`, rels files, and `[Content_Types].xml` overrides. exceljs can't write charts, so do this in `xlsxPostProcess.ts` on the zipped output (fflate is available). Preserve existing drawings (images) if any.
2. Read them in `fastXlsx.ts`: follow sheet rels → drawing → chart rels → chart XML and rebuild the app's chart objects (type, series refs, title, legend, anchor position/size).
3. Round-trip tests per chart type, plus a test that a file with charts made in real Excel (add a small fixture under `samples/`) loads its charts. Verify the saved file opens in Excel without a repair prompt.

You are working on "cellbook", a desktop Excel clone (Electron + React 19 + TypeScript + HyperFormula) at https://github.com/dhruv4740/cellbook. Another AI (Claude) reviews and merges your branch, so your work must end up pushed on your own branch. Goal of the app: look and behave EXACTLY like Microsoft Excel for Microsoft 365 on Windows, and stay fast on very large files (hundreds of thousands of rows).

## Setup

```bash
git clone https://github.com/dhruv4740/cellbook   # or: cd cellbook && git fetch origin
cd cellbook
git checkout claude/optimistic-einstein-x3mu90 && git pull
git checkout -b gemini/dv-save
npm install
```
Base is `claude/optimistic-einstein-x3mu90` (NOT main; it holds the latest merged work). Run with `npm run web` or `npm run dev`. Before pushing, `npm run typecheck` and `npm test` must pass and you must check the feature by hand. Don't commit `samples/sample.xlsx` (`git restore samples/sample.xlsx`). All edits go through store methods (`model/store.ts`) so they are undoable. Keep edits outside your task minimal. Put pure logic in `.ts` files with vitest tests next to them.

.xlsx files: reading uses the FAST reader `src/renderer/src/io/fastXlsx.ts` (hand-written XML scanner, runs in a worker; see how `parseConditionalFormats`/`parseDxfs` were added there). The exceljs reader in `io/xlsx.ts` (`readXlsxExceljs`) is legacy; writing uses exceljs in `writeXlsx` in `io/xlsx.ts` plus `io/xlsxPostProcess.ts` for raw XML patches exceljs can't do. Any new saved feature needs: write + fast read + a round-trip test in `io/xlsx.test.ts` (write then `readXlsx`) + it must open correctly in real Excel (valid OOXML).

When done: commit, `git push -u origin gemini/dv-save`, and reply with a short summary and anything unfinished.

## Task: Save and load Data Validation in .xlsx

Data validation rules live in `SheetLayout.dataValidations` (`DataValidationRule` in `model/types.ts`) and work in the app, but are lost on save.
1. Write them in `writeXlsx` as `<dataValidations>` / `<dataValidation type operator allowBlank showDropDown showInputMessage showErrorMessage errorStyle promptTitle prompt errorTitle error sqref><formula1/><formula2/></dataValidation>`. Note Excel's inverted flag: `showDropDown="1"` means the in-cell dropdown is HIDDEN. List sources are either `"a,b,c"` (quoted, comma list) or a range/formula. Use exceljs `ws.dataValidations.add` if it produces correct XML, otherwise patch in `xlsxPostProcess.ts`.
2. Read them in `fastXlsx.ts` (parse `<dataValidations>` in the sheet XML, multi-range `sqref`, decode entities).
3. Round-trip tests covering every Allow type, all three error styles, input/error messages, list with literal items and with a range source, dropdown on/off.

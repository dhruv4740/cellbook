# Cellbook vs real Excel — results

Machine: Windows 11, Microsoft Excel (desktop), region English (India) (`en-IN`, digit grouping 3;2, day-first dates).
Branch: `claude/optimistic-einstein-x3mu90`.

## 1. Formula correctness vs Excel

Method: `make_workbooks.py` → `cellbook_values.ts` → `scripts/excel-oracle/excel_compare.py` (opens `out/formulas.xlsx` in Excel over COM, `CalculateFull`, reads `Value2`/`Text` of A10:A170, compares; numbers within 1e-9 relative, errors by name).

Result: **155 formulas, 153 match.**

| Row | Formula | Excel | Cellbook | Status |
|---|---|---|---|---|
| 94 | `=TEXT(1234567.891,"#,##0.00")` | `12,34,567.89` | `1,234,567.89` | **Fixed** (see below) |
| 41 | `=1E+15+1` | `1000000000000001` | `#ERROR!` | Known unsupported: HyperFormula rejects the literal `1E+15` |
| 162 | `=INDIRECT("C"&2)` | `10` | `#NAME?` | Known unsupported: INDIRECT |

Fix: Excel formats `#,##0` with the PC's digit grouping, so on an English (India) PC 1234567 displays as `12,34,567` both in cells and in `TEXT()`. Cellbook always grouped by thousands. `engine/locale.ts` now detects lakh grouping from the OS (`LAKH_GROUPING`) and `format/numfmt.ts` uses it. Test: `numfmt.test.ts` "digit grouping follows the PC region".

Also fixed: two existing tests (`store.test.ts` "parses US dates", `numfmt.test.ts` "1/31/2025") failed on any day-first PC because they assumed US date order. Tests now pin the locale with `setDayFirst(false)`, and a new test checks day-first entry (`3/9/2026` = 3 Sept 2026). No expectation was changed.

Not checked by the sheet: `FIXED()`/`DOLLAR()` in `extraFunctions.ts` hard-code `en-US` grouping; not covered by the formula list.

## 2. Files saved by Cellbook, opened in Excel

Method: `npm run build`; `node scripts/excel-oracle/roundtrip.mjs <files>` (copies each file to `out/rt/`, opens it in the Electron app via `file:openPath`, presses Ctrl+S, waits for the file to change); `python scripts/excel-oracle/excel_check.py <files>` (opens in Excel over COM, watches `%TEMP%\error*.xml` and the window caption for a repair, `CalculateFull`, compares every formula's Excel result with the value cached in the file via openpyxl `data_only=True`, lists names/notes/links/CF/filters, then SaveAs `excel-<name>.xlsx`); then `compare_workbooks.py <cellbook-saved> <excel-resaved>`.

Needs `npm i --no-save playwright-core` (not in package.json). On this machine `npm ci` did not download the Electron binary; `node node_modules/electron/install.js` fixed it.

| File | Repair? | #NAME?/#REF! | Excel result ≠ cached | compare_workbooks |
|---|---|---|---|---|
| 1-saved-by-cellbook-features.xlsx | No | None | 0 | Only cosmetic (below) |
| 2-saved-by-cellbook-new-functions.xlsx | No | A13 `TEXTBEFORE` | 1 (A13) | 0 diffs |
| out/stress.xlsx | No | None | 0 | Only cosmetic (below) |

- Named range `TaxRate` (=Rates!$B$2) is in Name Manager; the note, link, conditional format, merges, freeze panes, hidden rows/cols and validation all survive (compare_workbooks shows no differences for them).
- `TEXTBEFORE` → `#NAME?` is **this Excel's limitation**, not Cellbook's: typing `=TEXTBEFORE("a-b","-")` or `=VSTACK(1,2)` into a new workbook in this Excel (16.0 build 20326, Office 2021-level) also gives `#NAME?`. Users on Excel 2021 or older will see `#NAME?` for TEXTBEFORE/TEXTAFTER/VSTACK etc. in files made with Cellbook. Cellbook's cached value (`a`) is correct.
- The `_xlfn.*` entries Excel lists in `Workbook.Names` for file 2 are Excel's own hidden placeholders; they are not in the file.
- Cosmetic diffs after an Excel re-save: Excel escapes `-` in number formats (`dd-mm-yyyy` → `dd\-mm\-yyyy`, same display) and adds a trailing `/` to `https://example.com`. Not real differences.
- Not done: opening visibly with DisplayAlerts on. Repair detection relied on the repair log in `%TEMP%` and the `[Repaired]` caption; neither appeared.

## 3. Files saved by Excel, opened in Cellbook

Method: `python scripts/excel-oracle/make_excel_workbook.py` builds `out/excel-built.xlsx` inside Excel over COM (committed copy: `samples/excel-made-features.xlsx`) with XLOOKUP, IFS, FILTER, UNIQUE, LET, a workbook name (`Rate`) and a sheet name (`Dispatch!LocalQty`), a table `Orders` with `=SUM(Orders[Qty])`, a note and a threaded comment, a list validation, a cell-rule CF, a 3-color scale, a data bar, a chart, merged M1:N2, frozen B2, hidden column J and row 9, the table filter Status = Open, dates `dd-mm-yyyy`, `#,##,##0.00`, Hindi text and leading-zero codes. Then `roundtrip.mjs` (open in the app, Ctrl+S), `compare_workbooks.py` and `excel_check.py`.

**Before the fixes this was the worst result of the whole check:** Cellbook could not save the file at all (Ctrl+S threw inside ExcelJS on the data bar, and nothing told the user), and once that was fixed, **Excel refused to open Cellbook's save** ("Open method of Workbooks class failed", even with repair). Found by bisecting the saved package in Excel. All fixed, each with a test:

| # | Bug | Effect | Fix |
|---|---|---|---|
| 1 | Data bars / color scales / icon sets written without thresholds | Ctrl+S failed on any sheet with one (including ones made in Cellbook) | `io/xlsx.ts` writes `cfvo`/colors; `io/fastXlsx.ts` now reads bar/scale colors and icon set names |
| 2 | Data validations and the chart `<drawing>` inserted *inside* the data-bar rule (they searched for the first `<extLst>`, which ExcelJS nests in the rule) | Excel refuses the file | `xlsxPostProcess.ts` `execTopLevel` |
| 3 | `<legacyDrawing>` (notes) written after the sheet `<extLst>` | Excel refuses the file | `fixLegacyDrawingOrder` |
| 4 | LET parameters saved without `_xlpm.` | Excel refuses the file | `engine/xlfn.ts` `addParamPrefixes` |
| 5 | `SUM(Orders[Qty])`: tables aren't modelled, the table was dropped but the reference kept | `#ERROR!` in Cellbook; Excel refuses the saved file | `engine/structuredRefs.ts`: structured refs become ranges on load (`SUM($C$2:$C$6)` = 80, as Excel) |
| 6 | Cellbook's `#ERROR!` cached as a cell error | Excel refuses the file | only Excel error codes are cached |
| 7 | A new function nested directly in a call (`ROWS(UNIQUE(..))`) saved without `_xlfn.` | `#NAME?` risk in Excel | `addFuturePrefixes` uses a look-behind |
| 8 | Excel's cached FILTER results (H9:H10) read as typed values | FILTER showed `#SPILL!` in Cellbook | reader drops values inside an array formula's `ref` |
| 9 | Spilled results saved as plain values, formula not flagged as a dynamic array | In Excel the FILTER no longer spilled; H9/H10 froze as `56`, `90` (leading zeros lost) | `store.toData` skips spill cells and records the spill range; writer emits `t="array" ref=…` + `cm="1"` + `xl/metadata.xml` like Excel |
| 10 | Frozen panes saved as `state="frozenSplit"` not recognised | Freeze lost | reader accepts `frozenSplit` |
| 11 | Table filter (lives in the table part) ignored | Rows 3 and 5 stayed hidden with no filter button | table's AutoFilter becomes the sheet AutoFilter |
| 12 | Threaded-comment fallback text uses `\r\n` in real Excel files | Comment text came back empty | `legacyNoteToThread` normalises line endings |

After the fixes: Excel opens Cellbook's save with no repair, no `#NAME?`/`#REF!`, every formula's Excel result equals Cellbook's cached value except LET (below), FILTER spills in Excel (`00012`, `00056`, `00090`, `HasSpill=True`), the comment text, the note, the filter, freeze, merges, hidden row/col, CF (3 rules) and the chart all survive. Cellbook shows `12,34,567.89` and `03-09-2026` like Excel.

Still open (not fixed):
- **LET is not implemented** in Cellbook's engine: `=LET(x,SUM(C2:C6),x*Rate)` shows `#NAME?` (Excel: 14.4). The formula now saves correctly, so Excel recalculates it to 14.4 on open, but the value cached in the file is `#NAME?` (Protected View would show that).
- **Tables are converted to plain ranges**: the ListObject (banding, header buttons, the name `Orders`) is not saved back. Formulas keep working because they are rewritten to `$C$2:$C$6`, but a row added below the old table is not included the way Excel would include it.
- Threaded comments come back as Cellbook comments with the author `tc={GUID}` (Excel keeps real authors in `xl/threadedComments`/`xl/persons`, which Cellbook doesn't read) and are saved as legacy notes.
- Hidden column J: Excel wrote width 0, Cellbook writes 13 with `hidden="1"`. Still hidden in Excel.
- Spilled cells (H9:H10) have no cached value in Cellbook's file; Excel fills them on open (the file has `fullCalcOnLoad`), but Protected View shows them blank.
- Not done: "any real .xlsx files the user points you to": none were provided.

## 4. Clipboard between Excel and Cellbook

**Not tested live.** Every process started from this Claude session runs inside a Windows job object that denies clipboard access (`OpenClipboard` → "Access is denied", also PowerShell `Set-Clipboard`/`Get-Clipboard`, also with the sandbox off). Nothing I run can read or write the Windows clipboard, so I didn't try to work around that. The first live attempt (`clipboard_check.py`) only showed Excel pasting its own copy.

What I did instead (`scripts/excel-oracle/clipboard_proxy.py`):
- Excel → Cellbook: Excel saves the block as Unicode text (the same tab-separated text it puts on the clipboard), and Cellbook's paste parses it.
- Cellbook → Excel: Cellbook's copy output (text + HTML) is captured; Excel opens the HTML with its HTML importer, the same one it uses for pasted HTML.

Bugs found and fixed (tests in `numfmt.test.ts`, `store.test.ts`, `commands.test.ts`):

| Bug | Before | After (= Excel on en-IN, checked with `FormulaLocal`) |
|---|---|---|
| **Typed or pasted `1,234.50` was stored as text** (HF has no thousands separator), so SUM ignored it. Same for `12,34,567.89` and `₹5,000` | text `"1,234.50"` | number 1234.5, format `#,##0.00`; `₹5,000` → 5000 with `₹ #,##0;[Red]₹ -#,##0`; `$5,000`, `1,23`, `12,345,67` stay text as in Excel |
| Copy: a cell with a line break wasn't quoted in the text | splits into 2 rows in text-only apps | `"Line1\nLine2"` like Excel |
| Copy: HTML had no line-break/text markers | Excel got `Line1 Line2` and `00012` → 12 | Excel gets `Line1⏎Line2` and `00012` (text) |
| Paste from Excel used only the text | a Text-formatted `00012` became 12 | reads Excel's HTML (`mso-number-format:"\@"` classes) and keeps it as text |

Also: `main` has no `readHTML` in Electron 44; added `clipboard:readHTML` through `clipboard.read()`. That IPC path is unit-tested only with a stub, not against the live clipboard.

**Please check by hand** (2 minutes, in the built app: `npm run build`, then `npx electron .`):
1. In Excel type in A1:E1: `1234.5` (format `#,##0.00`), `3/9/2026`, `a, b, c`, `Line1` Alt+Enter `Line2`, and `00012` in a Text-formatted cell. Copy, paste into Cellbook: expect 1234.5 as a number, 3 Sept 2026, one cell with two lines, `00012` kept.
2. Copy those cells in Cellbook, paste into Excel: same expectations.
3. Copy a cell in Cellbook, then copy a different cell in Excel, then Ctrl+V in Cellbook: it must paste the Excel cell. (Code reads the system text and drops Cellbook's own copy when it differs.)

Risk I could not rule out: if Cellbook's write to the Windows clipboard fails (clipboard locked by another app), its own Ctrl+C → Ctrl+V pastes nothing, because the system text then differs from the copy. Excel still pastes in that case.

## 5. Windows desktop behaviour

| Check | Result |
|---|---|
| `npm run dist` | Builds `release\Cellbook Setup 0.1.0.exe` (149 MB). Warning only: `author` missing in package.json. |
| Install the setup | **Not done.** Your installed Cellbook (`%LOCALAPPDATA%\Programs\Cellbook`, started 16:59) was running the whole time, and installing would have closed it. Run the setup yourself when it's closed. |
| Open by double-click | `.xlsx` is associated with Excel on this PC (`UserChoice` = `Excel.Sheet.12`), so a double-click opens Excel. Tested the packaged `release\win-unpacked\Cellbook.exe` started with a file path (what Explorer passes for "Open with Cellbook") in an isolated profile: the file opens in its own window. Launching it again with a second file hands that file to the running instance, and a window titled with the new file appears. The second part is inferred from window titles only. |
| Date entry follows the region | **Was broken, fixed.** Electron gives the page the Windows *display language* (en-US here), not the *region* (en-IN), so `3/9/2026` typed in the app was 9 March and digit grouping was western. Main now passes the region (`--cellbook-region=en-IN`, read by `engine/locale.ts` through the preload). Verified in the built app: `3/9/2026` → 46268 = 3 Sept 2026, shown `03-09-2026`. (Using Chromium's `--lang=en-IN` doesn't work: Chromium maps it to en-GB, which gets dates right but not lakh grouping.) |
| Ctrl+S while Excel has the file open | Correct: Cellbook shows "Your changes couldn't be saved to 'desktop-check.xlsx'. The file may be open in another program…", the file on disk is byte-identical, the edit is still in the app, and the workbook stays marked unsaved. No temp file left behind (only Excel's own `~$` lock file). |
| AutoSave | On: the file is saved within 6 s of an edit and the workbook is clean. Off: the file isn't touched for 6 s after an edit. The label shows On/Off. |
| Crash recovery | Correct: typed into a new Book1, waited for the 1-minute AutoRecover backup, killed the process tree (`taskkill /T /F`, like Task Manager's End task), relaunched. It shows "Cellbook closed unexpectedly last time. Your unsaved work in 'Book1' was recovered." and opens `Book1 (Recovered)` holding the edit. Edits made less than a minute before a crash are not in the backup (Excel's default interval is 10 minutes). |

Scripts: `desktop_check.py` / `desktop_check.mjs`.

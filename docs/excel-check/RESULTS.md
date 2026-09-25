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

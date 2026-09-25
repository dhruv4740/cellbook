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

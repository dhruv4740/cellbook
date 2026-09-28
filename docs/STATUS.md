# Status

**Scope: only the Windows desktop (Electron) app matters.** Test and verify there; the web build (`npm run web`) is not a target.

Checks on the latest commit: typecheck clean, 350 unit tests pass (also with `CELLBOOK_FAST_WRITE_MIN=0`, which forces the
fast file writer on every workbook), `npm run build` passes. Excel-compatibility results from the Windows + real Excel run
are in `docs/excel-check/RESULTS.md`; the oracle scripts in `scripts/excel-oracle/` drive the installed Excel over COM.

## Works like Excel

- **Files:** open/save .xlsx, .xlsm, .csv; safe save (temp file + swap); AutoSave; AutoRecover after a crash.
  Round-trips values, formulas, styles, number formats, merges, freeze panes, hidden rows/columns/sheets,
  notes, threaded comments, links, names, conditional formatting, data validation, filters, charts, page setup,
  headers/footers, print titles, page breaks, sheet/workbook protection (with Excel's password hashes), row/column
  groups, calculation mode.
- **Formulas:** 153/155 everyday formulas match real Excel (see RESULTS.md). SUBTOTAL/AGGREGATE honour filters and
  skip nested subtotals. COUNTIF(S), SUMIF(S), AVERAGEIF(S), MAXIFS, MINIFS, VLOOKUP, HLOOKUP, MATCH, XLOOKUP and INDEX
  run on a hash index (83+35 cases checked against real Excel, 0 differences). LET and INDIRECT (A1 text references) work.
- **Home:** clipboard (Paste options, Paste Special), font, Fill/Font colour palettes, borders, alignment, orientation,
  wrap, merge, number formats (₹ and dd-mm-yyyy on an English (India) PC), Increase/Decrease Decimal, Conditional
  Formatting, Format as Table, Cell Styles, Insert/Delete/Format, AutoSum, Fill, Clear, Sort & Filter, Find & Select.
- **Insert:** PivotTable, Table, charts, Link (Ctrl+K), Comment, Header & Footer (opens Page Setup).
- **Page Layout / Print:** Page Setup dialog (Page, Margins, Header/Footer with the custom editor, Sheet tab: print
  area, title rows/columns, gridlines, headings, black and white, draft, comments, error display, page order),
  margins presets, orientation, all paper sizes, breaks (insert/remove/reset), scale to fit width/height,
  a real pagination engine (page counts match Excel), print preview in File > Print, PDF export.
- **Formulas tab:** function menus, Name Manager, Show Formulas, Calculation Options (Automatic / Manual), F9,
  Shift+F9, "Calculate" in the status bar while manual results are stale.
- **Data:** Sort, Filter (Clear/Reapply), Advanced Filter (in place or copy, criteria ranges, unique only), Text to
  Columns, Flash Fill (Ctrl+E), Remove Duplicates, Data Validation, Consolidate (by position or labels), Goal Seek,
  Group/Ungroup with the outline gutter and level buttons, Subtotal, Refresh All, From Text/CSV.
- **Review / View:** protect sheet/workbook, comments, notes, gridlines, headings, formula bar, zoom, freeze panes,
  New Window, Arrange All, Switch Windows.
- **PivotTable:** up to two Rows fields, one Columns field, Sum/Count/Average/Max/Min, grand totals; refresh from the
  right-click menu or Data > Refresh All. Saved to .xlsx as values (not refreshable after reopening).
- **Motion:** menus, flyouts, dialogs, the File view and the ribbon fade/slide in like Excel's Fluent UI; hover
  feedback on buttons. All disabled by the Windows "reduce motion" setting.

## Performance (real despatch workbook scaled up, 21 columns)

| Rows | Open (parse + load) | Save | Sort | Insert rows / undo | 5,000 lookups |
|---|---|---|---|---|---|
| 521 | 0.3 s | 0.5 s | 0.2 s | 0.1 s | 0.6 s |
| 5,000 | 1.1 s | 0.7 s | 0.9 s | 0.5 s | 0.6 s |
| 50,000 | 6 s | 5 s | 8 s | 2 s | 0.7 s |

Sorting and removing duplicates are bound by HyperFormula's per-cell cost (about 1M cells rewritten). Scripts:
`scripts/perf-private.ts`, `perf-sort.ts`, `perf-write.ts`, `perf-lookup.ts`, `ui-perf.mjs` (see the top of each file).

## Not available yet (the button says so when clicked)

Pictures, shapes, icons, SmartArt, text boxes, WordArt, sparklines, slicers, timelines, maps, 3D Map, PivotChart;
Scenario Manager and Data Table (What-If); Forecast Sheet; spelling, thesaurus, translate; trace precedents/dependents
arrows, error checking, Watch Window; Page Break Preview and Page Layout view; split window; custom views; macros (VBA
is kept in .xlsm files but not run); themes; sheet background; Get Data from the web; R1C1 references in INDIRECT.

## Known limitations

- Excel tables load as ordinary ranges (formulas keep working); Format as Table is formatting + filter, not a table object.
- Manual calculation keeps formula results stale like Excel, but the engine still recalculates in the background, so it
  does not make edits faster.
- INDIRECT returns values only (no reference semantics, e.g. inside OFFSET or as a range argument to ROW).
- Threaded-comment authors from Excel show as `tc={GUID}`.
- Saving a very large workbook still runs on the UI thread (about 5 s at 50,000 rows).

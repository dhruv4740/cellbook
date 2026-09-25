# Status

**Scope: only the Windows desktop (Electron) app matters.** Test and verify there; the web build (`npm run web`) is not a target.

Checks on the latest commit: typecheck clean, 290/290 unit tests pass, `npm run build` passes, and the
features below were exercised in the built desktop app with real mouse/keyboard input (Playwright).
Excel-compatibility results from the Windows + real Excel run are in `docs/excel-check/RESULTS.md`.

## Works like Excel

- **Files:** open/save .xlsx, .xlsm, .csv; safe save (temp file + swap); AutoSave; AutoRecover after a crash.
  Round-trips values, formulas, styles, number formats, merges, freeze panes, hidden rows/columns/sheets,
  notes, threaded comments, links, names, conditional formatting, data validation, filters, charts, page setup.
- **Formulas:** 153/155 everyday formulas match real Excel (see RESULTS.md). SUBTOTAL/AGGREGATE honour
  filters and skip nested subtotals.
- **Home:** clipboard (Paste options, Paste Special), font, Fill/Font colour palettes, borders (edge presets
  apply to the selection edge), alignment, orientation, wrap, merge (Merge & Center toggles), number formats
  (₹ and dd-mm-yyyy on an English (India) PC), Increase/Decrease Decimal (keeps currency and colours),
  Conditional Formatting, Format as Table (banded style + filter), Cell Styles, Insert/Delete/Format
  (row height, column width, hide/unhide rows, columns and sheets, rename, move or copy, tab colour),
  AutoSum, Fill (Down/Right/Up/Left/Series), Clear (all, formats, contents, notes, links), Sort & Filter, Find & Select.
- **Insert:** PivotTable (see below), Table, charts, Link (Ctrl+K), Comment.
- **Page Layout / Print:** margins, orientation, paper size, print area, fit all columns on one page,
  print gridlines/headings. File > Print (Ctrl+P) prints the sheet through the Windows print dialog; Save as PDF.
- **Formulas tab:** function menus by category, Name Manager / Define Name / Use in Formula / Create from
  Selection, Show Formulas, Calculate Now (F9).
- **Data:** Sort, Filter (Clear/Reapply), Text to Columns, Remove Duplicates, Data Validation, Subtotal, Refresh All, From Text/CSV.
- **Review / View:** comments, notes (show/hide, show all), gridlines, headings, formula bar, zoom,
  zoom to selection, freeze panes, New Window, Arrange All, Switch Windows.
- **PivotTable:** up to two Rows fields (nested, with subtotals), one Columns field, Values with
  Sum/Count/Average/Max/Min, grand totals; Refresh from the right-click menu or Data > Refresh All.
  Saved to .xlsx as values: after reopening the file the numbers are there but it can't be refreshed.

## Not available yet (the button says so when clicked)

Pictures, shapes, icons, SmartArt, text boxes, WordArt, sparklines, slicers, timelines, maps, 3D Map,
PivotChart; Protect Sheet/Workbook; Group/Ungroup (outline); Advanced Filter; Flash Fill; Consolidate;
What-If Analysis; Forecast Sheet; spelling, thesaurus, translate; trace precedents/dependents arrows,
error checking, Watch Window; manual calculation mode; Page Break Preview; split window; custom views; macros (VBA is kept in .xlsm files but not run);
themes; headers and footers; print titles; Get Data from the web.

## Known limitations

- Excel tables load as ordinary ranges (formulas keep working); Format as Table is formatting + filter, not a table object.
- LET, INDIRECT and a literal `1E+15` inside a formula are not supported.
- Threaded-comment authors from Excel show as `tc={GUID}`.
- The Windows clipboard between Excel and Cellbook still needs one manual check (RESULTS.md §4).

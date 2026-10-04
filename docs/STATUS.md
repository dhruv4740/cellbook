# Status

**Scope: only the Windows desktop (Electron) app matters.** Test and verify there; the web build (`npm run web`) is not a target.

Checks on the latest commit: typecheck clean, 373 unit tests pass (also with `CELLBOOK_FAST_WRITE_MIN=0`, which forces the
fast file writer on every workbook), `npm run build` passes. Excel-compatibility results from the Windows + real Excel run
are in `docs/excel-check/RESULTS.md`; the oracle scripts in `scripts/excel-oracle/` drive the installed Excel over COM.

## Works like Excel

- **Files:** open/save .xlsx, .xlsm, .csv; safe save (temp file + swap); AutoSave; AutoRecover after a crash.
  Round-trips values, formulas, styles, number formats, merges, freeze panes, hidden rows/columns/sheets,
  notes, threaded comments, links, names, conditional formatting, data validation, filters, charts, page setup,
  headers/footers, print titles, page breaks, sheet/workbook protection (with Excel's password hashes), row/column
  groups, calculation mode.
- **Formulas:** 153/155 everyday formulas match real Excel (see RESULTS.md); 148 lookup / statistics cases match with 0 differences
  (`scripts/excel-oracle/func_oracle.py`). Also XMATCH, TAKE, DROP, CHOOSECOLS/ROWS, TOCOL/TOROW, SORTBY, MUNIT, MINVERSE, MDETERM,
  FREQUENCY, TREND, FORECAST(.LINEAR), INTERCEPT, PERCENTRANK, NUMBERVALUE, TYPE, VALUETOTEXT (the spilling ones need literal sizes;
  not yet: TEXTSPLIT, WRAPROWS/COLS, EXPAND, RANDARRAY, LAMBDA/MAP/REDUCE/SCAN/BYROW/BYCOL, CELL, INFO). SUBTOTAL/AGGREGATE honour filters and
  skip nested subtotals. COUNTIF(S), SUMIF(S), AVERAGEIF(S), MAXIFS, MINIFS, VLOOKUP, HLOOKUP, MATCH, XLOOKUP and INDEX
  run on a hash index (83+35 cases checked against real Excel, 0 differences). LET and INDIRECT (A1 text references) work.
- **Home:** clipboard (Paste options, Paste Special), font, Fill/Font colour palettes, borders, alignment, orientation,
  wrap, merge, number formats (₹ and dd-mm-yyyy on an English (India) PC), Increase/Decrease Decimal, Conditional
  Formatting, Format as Table, Cell Styles, Insert/Delete/Format, AutoSum, Fill, Clear, Sort & Filter, Find & Select.
- **Insert:** PivotTable, Table, charts, Pictures (file, or pasted screenshot; move/resize; saved in .xlsx), Shapes and Text Box
  (rectangles, oval, triangle, diamond, arrows, lines; edit text by double-click; saved in .xlsx), Sparklines (line, column,
  win/loss; saved as Excel's x14 extension), Symbol, Link (Ctrl+K), Comment, Header & Footer (opens Page Setup).
- **Page Layout / Print:** Page Setup dialog (Page, Margins, Header/Footer with the custom editor, Sheet tab: print
  area, title rows/columns, gridlines, headings, black and white, draft, comments, error display, page order),
  margins presets, orientation, all paper sizes, breaks (insert/remove/reset), scale to fit width/height,
  a real pagination engine (page counts match Excel), print preview in File > Print, PDF export.
- **Formulas tab:** function menus, Name Manager, Show Formulas, Calculation Options (Automatic / Manual), F9,
  Shift+F9, "Calculate" in the status bar while manual results are stale.
- **Data:** Sort, Filter (Clear/Reapply), Advanced Filter (in place or copy, criteria ranges, unique only), Text to
  Columns, Flash Fill (Ctrl+E), Remove Duplicates, Data Validation, Consolidate (by position or labels), Goal Seek,
  Group/Ungroup with the outline gutter and level buttons, Subtotal, Refresh All, From Text/CSV, Circle Invalid Data,
  Scenario Manager (saved in .xlsx), Data Table (writes values; not a live TABLE()).
- **Review / View:** Spelling (F7, offline English dictionary), Check Accessibility (basic), protect sheet/workbook,
  comments, notes, gridlines, headings, formula bar, zoom, freeze panes, Page Break Preview, New Window, Arrange All,
  Switch Windows.
- **Auditing:** Trace Precedents / Dependents arrows (one level per click), Error Checking with Excel's green triangles
  (number stored as text, formula errors), Watch Window, Evaluate Formula (no Step In / Step Out), Apply Names.
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

Icons, SmartArt, WordArt, slicers, timelines, maps, 3D Map, PivotChart; Forecast Sheet; thesaurus, translate; split
window; Page Layout view; custom views; macros (VBA is kept in .xlsm files but not run); themes; sheet background; Get Data
from the web; R1C1 references in INDIRECT. Shapes other than the basic set above, and shape formatting beyond fill,
outline and text, are not editable.

## Known limitations

- Excel tables are kept on save (name, range, style, banding, filter) and painted like Excel's built-in styles (approximate colours); formulas
  use plain ranges instead of structured references. Format as Table in Cellbook is still formatting + filter, not a table object.
- Large pictures live inside the sheet layout as base64, so every layout change (and each undo step) copies them; many big images make
  layout edits slower and use more memory.
- Manual calculation keeps formula results stale like Excel, but the engine still recalculates in the background, so it
  does not make edits faster.
- INDIRECT returns values only (no reference semantics, e.g. inside OFFSET or as a range argument to ROW).
- Threaded-comment authors from Excel show as `tc={GUID}`.
- Saving a very large workbook still runs on the UI thread (about 5 s at 50,000 rows).

## UI crawl (real app, every ribbon control + context menus)

`scripts/excel-oracle/ui_crawl.mjs <Tab>` clicks every ribbon control and dropdown item; `ui_context.mjs` does the cell, row, column and
sheet-tab right-click menus. Latest run: 0 renderer errors across all tabs and menus (Home 62, Insert 31, Draw 17, Formulas 26,
Data 32, Review 13, View 14, Help 7 controls). Added since: Forecast Sheet, Arrange (z-order, align, group, rotate, Selection Pane),
Select Objects, Copy as Picture, Draw/ink (pens, eraser, lasso, Ink to Shape, replay, Hide Ink), WordArt/Signature/Equation galleries,
Icons, Allow Edit Ranges, New Table Style, Custom Views, Themes/Colors/Fonts/Effects, Help links.

Still "isn't available" on purpose: Power Query (Get Data, From Web, Recent Sources, Connections, Data Model, Relationships), VBA
(Macros, View Code), Object, Maps, Stock/Online Pictures, 3D Models, Screenshot, SmartArt, Slicer, Timeline, Hierarchy/Statistic/
Waterfall/Funnel/Stock/Surface/Radar/Combo charts, Thesaurus, Translate, Smart Lookup, Quick Analysis, Show Changes, Unshare Workbook,
Split, Hide window, Ink to Math, Add-ins store, Sheet Background.
Simplifications: ink saved as freeform shapes (not InkML); custom table styles in localStorage; Custom Views session-only;
themes recolour new content only; WordArt/Equation are styled text boxes.

Also added after the first crawl: Quick Analysis, Page Layout Background (session only, not saved to .xlsx), View Hide/Unhide window,
Combo and Waterfall charts (saved as bar+line / stacked column), AutoFilter Date Filters, negative values on column/bar/line/area/combo
charts (value axis spans min..max). Fixed: Text/Number custom-filter dialogs closed instantly under the menu scrim; pen strokes moved the
cell selection. Removed from the "on purpose" list: Quick Analysis, Hide window, Combo, Waterfall, Background.
Still open: Split, Slicer, Timeline, SmartArt, Sort/Filter by Color, Sheet View in the filter menu, scatter charts with negatives.

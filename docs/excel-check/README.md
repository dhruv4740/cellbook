# Check in real Excel (5 minutes)

Both files were saved by Cellbook. Open each one in Microsoft Excel on Windows and note anything that looks wrong.

## 1-saved-by-cellbook-features.xlsx

1. **Opening:** Excel must open it without the message *"We found a problem with some content… Do you want us to try to recover?"*
2. **Sheet "Dispatch Data":**
   - Header row 1 is white bold text on dark blue, and the top-left pane is frozen at B2.
   - Column A shows codes with leading zeros: `00002`, `00003`, …
   - Column B has some Hindi names (`माल 5 ट्रक`), and B3 wraps onto several lines.
   - M1:O2 is one merged cell reading "Merged header".
   - Column I is hidden and row 10 is hidden.
   - A1 has a red note triangle, and hovering shows "Code has leading zeros".
   - B4 is a clickable link to example.com.
   - Q2:Q20 has a dropdown offering Yes / No / Maybe.
   - Column C values above 100 are pink with dark red text (conditional formatting).
3. **Formulas in column K**, each cell showing a number and not `#NAME?`:
   - K2 `=SUM(E2:E51)` shows **63345** (bold, red, italic, underlined).
   - K5 `=TaxRate*K2` shows **74747.1**. This uses the named range TaxRate; it should also appear in Formulas > Name Manager.
   - K6 `=VLOOKUP("00010",A:E,5,FALSE)` shows **405**.
   - K7 `=COUNTIF(B:B,"माल*")` shows **10**.
4. **Protected View:** email the file to yourself, then open the attachment. Before you click "Enable Editing", column K must already show the same numbers, not empty or 0.

## 2-saved-by-cellbook-new-functions.xlsx

Cells A2:A13 use XLOOKUP, IFS, CONCAT, TEXTJOIN, MAXIFS, IFNA, SWITCH, DAYS, AGGREGATE, STDEV.S, UNIQUE and TEXTBEFORE.
None of them may show `#NAME?`. Expected values, top to bottom:
`5, y, ab, 5,5, 5, 1, a, 9, 5, #DIV/0!, 5, a`

(`#DIV/0!` in A11 is correct: the standard deviation of a single number.)

## Also try, with your real files

- Open a real dispatch file in Cellbook, change a few cells, save it, then open it in Excel. Check that nothing moved, lost its colour or changed value.
- The other way round: save a file in Excel that uses XLOOKUP, IFS or named ranges, and open it in Cellbook.

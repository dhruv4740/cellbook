# Status (milestone-3)

Merged: Format Cells polish, Conditional Formatting, Data tools (Sort, Remove Duplicates, Text to Columns, Data Validation) from Gemini branches.
Fixed: fast .xlsx reader now reads conditional formatting (rules + dxf styles).
Checks: typecheck clean, 212/212 tests pass, `npm run build` passes.

## Known gaps
- Data validation rules are not yet written to / read from .xlsx (in-app only).
- Charts are not saved to .xlsx.
- View > New Window not built.

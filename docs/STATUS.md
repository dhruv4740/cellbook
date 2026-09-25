# Status (milestone-3)

Merged: Format Cells polish, Conditional Formatting, Data tools (Sort, Remove Duplicates, Text to Columns, Data Validation) from Gemini branches.
Fixed: fast .xlsx reader now reads conditional formatting (rules + dxf styles).
Checks: typecheck clean, 212/212 tests pass, `npm run build` passes.

## Known gaps
- Data validation not saved to .xlsx → docs/gemini-prompts/4
- Charts not saved to .xlsx → docs/gemini-prompts/5
- View > New Window → docs/gemini-prompts/6

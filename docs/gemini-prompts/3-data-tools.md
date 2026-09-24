You are working on "cellbook", a desktop Excel clone (Electron + React 19 + TypeScript + HyperFormula) at https://github.com/dhruv4740/cellbook. Another AI (Claude) works on the same repo from the cloud and will pull your branch to review and merge it, so your work must end up pushed on your own branch. The goal of the whole app: look and behave EXACTLY like Microsoft Excel for Microsoft 365 on Windows (light theme), with Excel's exact names, labels and layouts.

## Setup

```bash
git clone https://github.com/dhruv4740/cellbook   # or: cd cellbook && git fetch origin
cd cellbook
git checkout main && git pull
git checkout -b gemini/data-tools
npm install
```
Run with `npm run web` (browser at http://localhost:5173) or `npm run dev` (desktop). Before pushing, `npm run typecheck` and `npm test` must pass, and you must open the app and check your feature works by hand. Don't commit changes to `samples/sample.xlsx` (tests regenerate it; `git restore samples/sample.xlsx`).

Read first: `src/renderer/src/model/types.ts`, `model/store.ts` (all edits must go through store methods so they are undoable: `setCells`, `applyStyle`, `updateLayout`, `transact`), `chrome/ui.ts`, `components/Dialog.tsx`, `components/Menu.tsx`, `dialogs/DialogHost.tsx` (register new dialogs here by `ui.dialog.kind`), `grid/commands.ts`. Keep edits to files outside your task minimal and additive. Put pure logic in `.ts` files with vitest tests next to them.

When done: commit, `git push -u origin gemini/data-tools`, and reply with a short summary and anything unfinished.

## Task: Data tab tools (dialogs + logic, no grid rendering changes)

Build these with Excel's exact dialogs, wire them to the Data tab buttons (`chrome/ribbon/OtherTabs.tsx`) and the Home > Sort & Filter menu (`chrome/ribbon/HomeTab.tsx`), and put logic in `data/*.ts` with vitest tests:
1. **Sort** (Data > Sort): the Sort dialog with "+ Add Level", "Delete Level", "Copy Level", up/down arrows, "Options..." (Case sensitive, Orientation: Sort top to bottom / Sort left to right), "My data has headers" checkbox (auto-detected like Excel), columns Column / Sort On (Cell Values, Cell Color, Font Color, Conditional Formatting Icon) / Order (A to Z, Z to A, Custom List...; Smallest to Largest for numbers). Sorting moves whole rows with their styles, keeps formulas' relative references correct (use `engine/shiftFormula.ts`), is one undo step, and follows Excel's order: numbers < text < logicals < errors < blanks (blanks always last). Also the quick "Sort A to Z" / "Sort Z to A" buttons and Excel's "Sort Warning" dialog (Expand the selection / Continue with the current selection) when the selection is next to other data. Replace the existing simple `sortSelection` in `grid/commands.ts` with your implementation (same signature).
2. **Remove Duplicates**: dialog with "Select All", "Unselect All", "My data has headers", column checklist; result message "N duplicate values found and removed; M unique values remain." (Excel's wording).
3. **Text to Columns**: the 3-step Convert Text to Columns Wizard (Delimited / Fixed width; delimiters Tab, Semicolon, Comma, Space, Other, Treat consecutive delimiters as one, Text qualifier; data preview; column data format General/Text/Date/Do not import; Destination).
4. **Data Validation** (Data > Data Validation): the dialog with Settings (Allow: Any value, Whole number, Decimal, List, Date, Time, Text length, Custom; Data: between/not between/equal to/...; Minimum/Maximum; Ignore blank; In-cell dropdown; Source), Input Message and Error Alert tabs (Stop / Warning / Information). Store rules in `SheetLayout.dataValidations` (additive to `model/types.ts`, changed via `store.updateLayout`). Enforce on commit: export `validateInput(store, sheet, r, c, raw): { ok: true } | { ok: false, style, title, message }` and call it from `commitEdit` in `chrome/ui.ts`, showing Excel's error box ("This value doesn't match the data validation restrictions defined for this cell." with Retry / Cancel for Stop). List validation shows the dropdown arrow button next to the active cell: render it as an absolutely positioned HTML button over the grid (read the active cell's rectangle from a small exported helper you add to `grid/Grid.tsx`) that opens a list; picking an item sets the cell.
5. **Flash Fill** is out of scope; leave the button inert.

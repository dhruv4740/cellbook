You are working on "cellbook", a desktop Excel clone (Electron + React 19 + TypeScript + HyperFormula) at https://github.com/dhruv4740/cellbook. Another AI (Claude) works on the same repo from the cloud and will pull your branch to review and merge it, so your work must end up pushed on your own branch. The goal of the whole app: look and behave EXACTLY like Microsoft Excel for Microsoft 365 on Windows (light theme), with Excel's exact names, labels and layouts.

## Setup

```bash
git clone https://github.com/dhruv4740/cellbook   # or: cd cellbook && git fetch origin
cd cellbook
git checkout main && git pull
git checkout -b gemini/format-cells-polish
npm install
```
Run with `npm run web` (browser at http://localhost:5173) or `npm run dev` (desktop). Before pushing, `npm run typecheck` and `npm test` must pass, and you must open the app and check your feature works by hand. Don't commit changes to `samples/sample.xlsx` (tests regenerate it; `git restore samples/sample.xlsx`).

Read first: `src/renderer/src/model/types.ts`, `model/store.ts` (all edits must go through store methods so they are undoable: `setCells`, `applyStyle`, `updateLayout`, `transact`), `chrome/ui.ts`, `components/Dialog.tsx`, `components/Menu.tsx`, `dialogs/DialogHost.tsx` (register new dialogs here by `ui.dialog.kind`), `grid/commands.ts`. Keep edits to files outside your task minimal and additive. Put pure logic in `.ts` files with vitest tests next to them.

When done: commit, `git push -u origin gemini/format-cells-polish`, and reply with a short summary and anything unfinished.

## Task: finish the Format Cells dialog (`src/renderer/src/dialogs/FormatCellsDialog.tsx/.css`)

It exists but is simplified. Make it match Excel exactly:
- **Alignment**: Horizontal list must be General, Left (Indent), Center, Right (Indent), Fill, Justify, Center Across Selection, Distributed (Indent); Vertical: Top, Center, Bottom, Justify, Distributed; add the "Justify distributed" checkbox; Text control must have Wrap text, Shrink to fit, **Merge cells** (merges/unmerges via `mergeCells` in grid/commands.ts); the Orientation box must be Excel's: a narrow vertical "Text" button on the left and a half-circle dial on the right with dots every 15°, a red-diamond handle and the word "Text" drawn along the current angle, draggable with the mouse, synced with the Degrees spinner (-90..90).
- **Border**: Line Style must be Excel's 2-column grid of 14 styles (None + 13), exactly one selected; Presets None, Outline, **Inside** (enabled only when several cells are selected); the preview has the 8 toggle buttons around it (top, middle horizontal, bottom, diagonal up on the left column; left, middle vertical, right, diagonal down along the bottom) with Excel's icons, and clicking the preview's edges toggles them; the preview shows "Text" samples in 1 or 4 quadrants depending on the selection; the hint sentence under it. Apply Inside borders by setting top/left borders on the inner cells (use `store.applyStyle` with a function patch).
- **Font**: Font list is a text box above a scrolling list, same for Font style and Size; Underline has None, Single, Double, Single Accounting, Double Accounting; Color is a dropdown with Automatic + the theme palette; Normal font checkbox; Effects: Strikethrough, Superscript, Subscript; Preview box ("AaBbCcYyZz" in the chosen font); the TrueType note.
- **Fill**: Background Color: "No Color" button, 10x6 theme grid, standard colors row, Fill Effects... and More Colors... buttons, Pattern Color and Pattern Style dropdowns, Sample.
- **Number**: every category's controls (Decimal places spinner, Use 1000 Separator, Symbol, Negative numbers list with red entries, Type lists, Locale, Custom Type box + list + Delete) and Excel's description sentence for each category. The patch applied on OK must only include what the user changed.
- OK applies to the whole selection as one undo step.

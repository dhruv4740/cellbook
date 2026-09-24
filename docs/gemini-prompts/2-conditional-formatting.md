You are working on "cellbook", a desktop Excel clone (Electron + React 19 + TypeScript + HyperFormula) at https://github.com/dhruv4740/cellbook. Another AI (Claude) works on the same repo from the cloud and will pull your branch to review and merge it, so your work must end up pushed on your own branch. The goal of the whole app: look and behave EXACTLY like Microsoft Excel for Microsoft 365 on Windows (light theme), with Excel's exact names, labels and layouts.

## Setup

```bash
git clone https://github.com/dhruv4740/cellbook   # or: cd cellbook && git fetch origin
cd cellbook
git checkout main && git pull
git checkout -b gemini/conditional-formatting
npm install
```
Run with `npm run web` (browser at http://localhost:5173) or `npm run dev` (desktop). Before pushing, `npm run typecheck` and `npm test` must pass, and you must open the app and check your feature works by hand. Don't commit changes to `samples/sample.xlsx` (tests regenerate it; `git restore samples/sample.xlsx`).

Read first: `src/renderer/src/model/types.ts`, `model/store.ts` (all edits must go through store methods so they are undoable: `setCells`, `applyStyle`, `updateLayout`, `transact`), `chrome/ui.ts`, `components/Dialog.tsx`, `components/Menu.tsx`, `dialogs/DialogHost.tsx` (register new dialogs here by `ui.dialog.kind`), `grid/commands.ts`. Keep edits to files outside your task minimal and additive. Put pure logic in `.ts` files with vitest tests next to them.

When done: commit, `git push -u origin gemini/conditional-formatting`, and reply with a short summary and anything unfinished.

## Task: Conditional Formatting (Home > Styles > Conditional Formatting)

1. Model: add `conditionalFormats?: CfRule[]` to `SheetLayout` in `model/types.ts` (additive; changes go through `store.updateLayout` so they undo). Rule types: cellIs (greater than, less than, between, equal to, not equal...), textContains, dateOccurring (Yesterday, Today, Tomorrow, In the last 7 days, Last/This/Next week, Last/This/Next month), duplicateValues/uniqueValues, top/bottom N and N%, above/below average, expression (formula, evaluated relative to the rule range's top-left via HyperFormula `store.hf.calculateFormula` with references shifted per cell; `engine/shiftFormula.ts` has a helper), dataBar, colorScale (2- and 3-color), iconSet (3 Arrows, 3 Traffic Lights, 3 Symbols, 4/5 Ratings...). Each has ranges, priority, stopIfTrue, and a format (font color, bold, italic, underline, strike, fill, border, number format) or visual settings.
2. Engine: `format/conditional.ts` with `evaluateCf(store, sheet, r, c): { style?: Partial<CellStyle>; dataBar?: { fraction: number; color: string; negative?: boolean }; colorScale?: string; icon?: { set: string; index: number } }`, cached per sheet and invalidated when `store.version` changes. Full vitest coverage.
3. Rendering: in `grid/Grid.tsx`, where each cell's style is resolved before drawing, merge `evaluateCf(...).style` over the cell style, and draw data bars (Excel's gradient fill bar with solid border), color-scale backgrounds and icon-set icons (draw them as small canvas shapes). Keep the grid change small and fast: evaluate only visible cells.
4. UI (Excel's exact menu): Highlight Cells Rules › (Greater Than..., Less Than..., Between..., Equal To..., Text that Contains..., A Date Occurring..., Duplicate Values..., More Rules...), Top/Bottom Rules › (Top 10 Items..., Top 10%..., Bottom 10 Items..., Bottom 10%..., Above Average..., Below Average..., More Rules...), Data Bars › (Gradient Fill and Solid Fill galleries, 6 colors each), Color Scales › (12-swatch gallery), Icon Sets › (Directional, Shapes, Indicators, Ratings galleries), separator, New Rule..., Clear Rules › (Clear Rules from Selected Cells, Clear Rules from Entire Sheet), Manage Rules.... Build the quick dialogs ("Greater Than": "Format cells that are GREATER THAN:" value box + "with" dropdown of presets: Light Red Fill with Dark Red Text, Yellow Fill with Dark Yellow Text, Green Fill with Dark Green Text, Light Red Fill, Red Text, Red Border, Custom Format...), the New Formatting Rule dialog (rule-type list + editor per type) and the Conditional Formatting Rules Manager (list with Rule/Format/Applies to/Stop If True, New/Edit/Delete/Duplicate, up/down priority). Live preview while the quick dialog is open, like Excel.
5. Wire the Home tab's Conditional Formatting button (`chrome/ribbon/HomeTab.tsx`) to this menu.
6. Save/load rules in `io/xlsx.ts` if exceljs supports it (it has `worksheet.addConditionalFormatting` and reads `worksheet.conditionalFormattings`); round-trip test.

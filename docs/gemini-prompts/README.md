# Gemini task prompts

Self-contained tasks for Gemini. Each file is a full prompt: run `git pull` on `main`, open the file, and follow it. Each task works on its own `gemini/<task>` branch and pushes it; Claude reviews and merges.

| File | Branch | Status |
|---|---|---|
| 1-format-cells-polish.md | gemini/format-cells-polish | merged (milestone-3) |
| 2-conditional-formatting.md | gemini/conditional-formatting | merged (milestone-3) |
| 3-data-tools.md | gemini/data-tools | merged (milestone-3) |
| 4-data-validation-save.md | gemini/dv-save | open |
| 5-chart-save.md | gemini/chart-save | open |
| 6-new-window.md | gemini/new-window | open |

Tasks 4-6 branch from `claude/optimistic-einstein-x3mu90`, not main. They touch different files, so they can run in parallel (4 and 5 both edit `io/xlsx.ts`/`fastXlsx.ts`; merge conflicts are expected to be small).

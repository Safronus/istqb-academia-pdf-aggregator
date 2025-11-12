# ISTQB Academia PDF Aggregator

## Version 0.10k — 2025-11-12
### Overview (selection stability + visuals)
- **Selection stability:** Selecting one row no longer leads to *selecting all* after the Sorted refresh.
  - During Sorted refresh and row hiding, the code now **captures current selection by file name**,
    updates the model **without noisy signals**, and **restores exactly the same selection** afterwards.
  - No auto-refresh hooks that would re-trigger selection are invoked during user interaction.
- **Sorted column appearance:** kept centered **check icon** and added **light‑gray background** for the whole column (unchanged from 0.10j).

### What changed technically
- `_overview_update_sorted_flags(...)` now:
  - caches **selected keys** (from “File name”),
  - blocks selection signals and repaints while updating,
  - sets data **only if changed**, and finally **restores selection** by keys.
- `_overview_apply_sorted_row_hiding(...)` now:
  - uses the same **selection preserve/restore** approach while toggling visibility,
  - only touches rows whose visibility actually changes.

### Other notes
- No changes to exports, context menu, filters, or proxy behaviour.
- Minimal‑change philosophy preserved; PySide6 only; macOS friendly.


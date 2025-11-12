# ISTQB Academia PDF Aggregator

## Version 0.10o — 2025-11-12
### Overview
- **Sorted column finally stable:** `rescan()` now builds the Overview with the **Sorted** column present
  and fills it right after repopulation using SHA‑256 matches against *Sorted PDFs*.
- **Selection stability:** `rescan()` now restores selection using the explicit **File name** column
  (not by assuming it's the last column). This prevents deselection when auto-rescan triggers.
- **Icons & visuals:** Wished Recognition icons and the Sorted icon are centered; Sorted column keeps a light‑gray background.
- **Board combobox:** two sections (present boards → separator → remaining known boards) intact.

No other behaviour changed; minimal‑change patch for macOS / PySide6.


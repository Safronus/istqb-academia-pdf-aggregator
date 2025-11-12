# ISTQB Academia PDF Aggregator — Release v0.10o
**Date:** 2025-11-13  •  **Platform:** macOS  •  **GUI:** PySide6  •  **Theme:** Dark

## Highlights
- **Overview**
  - New **Sorted** column with **SHA‑256** matching vs. *Sorted PDFs* (recursive, case‑insensitive folder discovery).
  - **Stable selection**: no more auto select/deselect during refresh or rescan; selection is preserved on rebuild.
  - **Board filter combobox with two sections**: first boards currently present in the table, separator, then the remaining known boards.
  - **Centered icons** in *Wished Recognitions* and **Sorted** columns; **Sorted** column has a light‑gray background.
  - **Board column de‑dup** (visual only): shows the board name only on the first occurrence per group (keeps data intact).
  - **Export (Overview)**: export **Selected rows** or **All rows** to **CSV / XLSX / TXT**. TXT uses a **titled + sectioned** layout with bullet sub‑sections.
  - **Context export to “Sorted PDFs”** from the table.

- **PDF Browser**
  - Vertical reflow: **file list on top**, **details below** (better room for long names).
  - **Detail panel** mirrors Overview fields (excludes *Known Board* and *Sorted Status* by request).
  - Window & sections **autosize** to fit content.

- **Sorted PDFs (DB) tab**
  - Same export options as in Overview (CSV / XLSX / TXT).

- **Board Contacts (v0.9+)**
  - New tab **Board Contacts** — per‑board contact directory (**Full Name + Email**).
  - JSON storage (git‑ignored). **Import from CSV** supported. Columns auto‑fit.

- **Recognized People List (v0.10+)**
  - New tab **Recognized People List** with **Add person** dialog.
  - Add from **Sorted PDFs DB** or manually; prefill **badge types** (Academia/Certified), **Recognition Date**, **Link to badge**, **Board**, **Full Name**, **Email**, **Address**.
  - **Duplicate check** when adding.
  - **Valid Until** = Recognition Date + **365 days**.
  - Row **coloring by validity**: green (>1 month left), yellow (<1 month left), red (expired).
  - Fulltext filter + three toggles: **Valid**, **Before expiry**, **Expired**.

- **Parser/Scanner**
  - **Application Type cleanup** — leading '/' stripped at parsing time.

## Fixes (selected)
- Fixed **auto-rescan** side effects: selection is preserved and does not “select all” or get cleared.
- **Sorted** status recalculation does not interfere with user selection.
- Board filter combobox **restored** to two-section behavior.
- Multiple stability & layout adjustments across tabs.

## Breaking Changes
- **None.** All changes are backward compatible and minimal-change by design.

## Known Issues
- None reported in this release. If a *Sorted PDFs* folder has a non-standard name, ensure it matches one of: `Sorted PDFs`, `sorted_pdfs`, or `sorted-pdfs` (case-insensitive), or place it under the project root/PDF root.

## Install / Run (macOS)
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# (optional) pip install openpyxl   # for XLSX export
python -m app
```

## Data Files (git-ignored)
- `board_contacts.json` — Board Contacts tab storage.
- `recognized_people.json` — Recognized People List storage.
- `Sorted PDFs/` — exported/processed PDFs.

## How to Use
- **Overview**: filter by Board (two-section combo) or search, select rows, right-click for **Export to Sorted PDFs** or use **Export…** to generate CSV/XLSX/TXT (selected or all).
- **PDF Browser**: navigate PDFs (top), see details (bottom). Double-click to open PDF.
- **Sorted PDFs**: edit record fields if needed; export DB using **Export…**.
- **Board Contacts**: import from CSV or edit inline; autosave to JSON.
- **Recognized People List**: add from DB or manual; filters and coloring help with validity management.

## Changelog (condensed)
- **0.10o**: Fixed rescan rebuild (includes Sorted column), improved selection restore by explicit File name column.
- **0.10n**: Moved helper methods out of nested scope; stable in‑place Sorted update.
- **0.10m**: Removed intrusive auto-refresh during selection; restored two-section Board combobox.
- **0.10j–0.10l**: Gray background for Sorted; selection guard; stability tweaks.
- **0.10g–0.10i**: Sorted detection via SHA‑256; centered icons.
- **0.10**: Recognized People List tab with add/edit/delete, validity logic, filters.
- **0.9–0.9a**: Board Contacts tab + CSV import & help.
- **0.8**: Window/sections size tuning and fitting.
- **0.7**: PDF Browser vertical layout and details fitting.
- **0.6a–0.6c**: Selection fixes; English warning; Application Type cleanup.

## Credits
- GUI: **PySide6**. Design: minimal-change + dark theme for macOS, HiDPI-ready.
- Thanks for feedback & testing!


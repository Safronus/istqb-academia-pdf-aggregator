# ISTQB Academia PDF Aggregator

**Version:** v0.10o  •  **Date:** 2025-11-13  •  **Platform:** macOS  •  **GUI:** PySide6

A desktop tool to scan, preview and manage ISTQB Academia Recognition PDFs. Features include Overview filtering, robust export, a PDF browser, a *Sorted PDFs* DB view, a Board Contacts directory, and a Recognized People List.

## Quick Start (macOS)
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# (optional) pip install openpyxl   # for XLSX export
python -m app
```

## Key Features
- **Overview**
  - *Board* filter combobox in two sections (present → separator → remaining known).
  - *Search* fulltext filter.
  - **Sorted** column based on **SHA‑256** matches vs. *Sorted PDFs* (recursive).
  - Context export to **Sorted PDFs**; **Export…** to **CSV / XLSX / TXT** (selected or all).
  - Centered icons in Wished & Sorted columns; Sorted has light‑gray background.
  - Board column shows repeated values only once (visual).

- **PDF Browser**
  - Vertical layout (files top, details bottom), detail panel mirrors Overview.
  - Double‑click to open PDF. Window/sections auto‑fit long names.

- **Sorted PDFs**
  - Same export dialog/format options as Overview.

- **Board Contacts**
  - Per‑board contact directory (Full Name + Email). CSV import. JSON storage (git‑ignored).

- **Recognized People List**
  - Add from DB or manual; duplicates check.
  - Recognition Date, Badge types, Link to badge; **Valid Until = +365 days**.
  - Row coloring by validity (green/yellow/red) and filters (Valid / Before expiry / Expired).

## Data & Storage
- `Sorted PDFs/` (folder): processed PDFs (git‑ignored).
- `board_contacts.json` and `recognized_people.json`: local JSONs (git‑ignored).

## Export Formats
- **CSV** — comma‑separated, headers included.
- **XLSX** — requires `openpyxl`.
- **TXT** — titled and sectioned layout with bullet sub‑sections.

## Notes
- Minimal-change philosophy: preserve existing behavior/design; PySide6 only; dark theme by default.
- HiDPI/Retina friendly.


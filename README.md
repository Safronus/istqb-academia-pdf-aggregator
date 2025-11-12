# ISTQB Academia PDF Aggregator

## Version 0.8c — 2025-11-12
### PDF Browser
- **Details panel:** hide fields **"Known Board"** and **"Sorted Status"** (labels remain instantiated for compatibility, only not shown in the form).

### Notes
- Minimal-change: no logic altered, no signal/slot changes.
- Other tabs (Overview, Sorted PDFs) unaffected.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

### Recent
- 0.8c — PDF Browser: hide "Known Board" and "Sorted Status" in details panel.
- 0.8b — Overview export: scope option (Selected rows vs All rows).
- 0.8 — Overview Board combobox grouped (present + separator + remaining).

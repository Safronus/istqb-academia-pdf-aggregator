# ISTQB Academia PDF Aggregator

## Version 0.7 — 2025-11-12
### What changed (PDF Browser)
- **Layout reflow:** In the *PDF Browser* tab the **details panel is now below** the file list (not at the right), so the file **Name** column gets more horizontal space.
- **Fit data comfortably:** The details panel uses wrapped labels/values and grows to fit; on first directory load the app **widens the window if needed** so both the Name list and details fit without truncation.
- **Minimal-change:** No changes to parsing, export, or Overview. All label names stay the same so `_update_detail_panel(...)` continues to fill all fields.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

### Release history
- **0.6c** — fix(scanner): remove leading '/' from *Application Type* (AcroForm) in `_parse_one`.
- **0.7** — *PDF Browser* layout reflow (details under list) + one-shot window widen to fit data.


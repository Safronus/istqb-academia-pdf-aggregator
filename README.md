# ISTQB Academia PDF Aggregator

## Version 0.10c — 2025-11-12
### Recognized People List
- **Add person…** now **splits multiple badges into separate records** (e.g., Academia **and** Certified -> two rows), inserted together (grouped visually).
- **Valid Until** is **auto-computed** as *Recognition Date + 1 year* on load, add and edit.
- **Row coloring** (applied when switching to the tab):
  - **Green**: validity > 1 month remaining
  - **Yellow**: ≤ 1 month remaining (still valid)
  - **Red**: expired (after *Valid Until*)
- **Edit…**: if you select both badges for a single-row edit, current row is updated for one badge and the **missing badge** is optionally inserted as another row (if not a duplicate).
- Column widths keep **auto-fitting** to content.

### Notes
- JSON persistence unchanged (`recognized_people.json`, git-ignored).
- No changes to other tabs.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

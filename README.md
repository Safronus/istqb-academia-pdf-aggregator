# ISTQB Academia PDF Aggregator

## Version 0.7e — 2025-11-12
**Sorted PDFs export = identical to Overview export**  
- Added Export dialog in Sorted PDFs with the **same options** as Overview:
  - choose **formats** (XLSX/CSV/TXT),
  - choose **Boards** (All or specific),
  - choose **Fields/columns** (same order/labels).
- Reuses the same export helpers: `_export_to_xlsx/_export_to_csv/_export_to_txt`.
- Data source: **DB** (`self.sorted_db.get(path)`), not UI fields.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

### Release History (recent)
- 0.7e — Sorted export dialog & options mirror Overview.
- 0.7d — Sorted export columns, filters, engines same as Overview.
- 0.7c — add `ed_sigdate` alias.
- 0.7b — add `ed_rec_acad/ed_rec_cert` + aliases.
- 0.7a — fix field names/aliases in Sorted builder.
- 0.7 — add Export… button in Sorted, basic export.
- 0.6c — scanner: strip leading '/' from Application Type.

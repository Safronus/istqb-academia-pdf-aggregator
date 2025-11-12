# ISTQB Academia PDF Aggregator

## Version 0.10g — 2025-11-12
### Overview — Sorted indicator by PDF hashes
- The **Sorted** column is now computed by **cryptographic hash match (SHA‑256)** between the source PDF and any PDF found under **“Sorted PDFs”** (recursive).
- If a row's file hash matches a hash in “Sorted PDFs”, the indicator shows **Yes**; otherwise it's empty.
- Hashing uses a **chunked reader** and a small in‑memory **cache** to avoid re‑hashing the same file repeatedly during a session.
- The indicator refreshes automatically when the Overview table rebuilds or after rescans.

### Notes
- No changes to existing UI/behaviour besides the Sorted indicator logic.
- Fallback filename matching is no longer used; only **hash equality** determines Sorted status.
- macOS HiDPI friendly; PySide6 only.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

# ISTQB Academia PDF Aggregator

## Version 0.10h — 2025-11-12
### Overview — robust 'Sorted' indicator
- The **Sorted** column is now computed using **SHA‑256** match against PDFs under any folder named **“Sorted PDFs”** (case-insensitive) under **repo root** or **pdf_root**.
- For each Overview row, the PDF path is inferred from:
  1) attached record path (if present),
  2) a direct path in the “File name” cell (if it's a path),
  3) a recursive search within **pdf_root** by file name.
- Added internal caches for file hashes and name→path resolution to avoid re-hashing.
- The logic is applied on table rebuild and rescans.

_No UI changes except the indicator value._

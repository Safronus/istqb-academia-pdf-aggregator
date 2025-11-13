# Release v0.11a — Parser & Scanner update
**Date:** 2025-11-13

## What changed
- **Parser (`pdf_parser.py`)**
  - Reads **AcroForm** fields first (if present), otherwise falls back to **text** patterns (`Key: Value` lines).
  - New keys populated (if present): `printed_name_title`, `istqb_receiving_board`, `istqb_date_received`, `istqb_valid_from`, `istqb_valid_to`.
  - Robust to label variants: supports **"Printed Name, Title"** and **"Name and title"**.

- **Scanner (`pdf_scanner.py`)**
  - `PdfRecord` dataclass extended with the fields above.
  - `_parse_one()` enriches the record with parser output.

## Backward compatibility
- Older DB entries or PDFs without these fields remain valid; missing values are empty/`None`.


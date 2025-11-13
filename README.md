# ISTQB Academia PDF Aggregator

**Version:** v0.11a  •  **Date:** 2025-11-13  •  **Platform:** macOS  •  **GUI:** PySide6

This version extends the **parser & scanner** to capture additional fields from the PDF:

- **Section 6 – Declaration and Consent**
  - `printed_name_title`
- **Section 7 – For ISTQB Academia Purpose Only**
  - `istqb_receiving_board`
  - `istqb_date_received`
  - `istqb_valid_from`
  - `istqb_valid_to`

These values are exposed across the app (Overview / PDF Browser / Sorted DB) via the same attribute names.

## Quick Start (macOS)
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```


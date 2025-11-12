# ISTQB Academia PDF Aggregator (PySide6)
**Version:** 0.6a — 2025-11-12

A macOS‑oriented **PySide6** GUI app (English UI) that scans PDF files inside the project’s **`PDF/`** folder (including subfolders named by ISTQB member boards), parses key fields from the *ISTQB Academia Recognition Program – Application Form*, and aggregates them into a tabular overview.

---

## Key Features
- **Overview** tab: consolidated table of parsed fields from all PDFs under `PDF/` (subfolders = **board names**). Columns are visually grouped (Application, Institution, Wished Recognitions, Contact details, Eligibility, Signature). “Academia/Certified Recognition” show **check/close** icons for `Yes/No`.  
- **PDF Browser** tab: file‑centric view; shows **all fields** for the selected PDF (including **Eligibility**).
- **Sorted PDFs** tab *(if the folder exists)*: mirrors **Overview**, but scans the separate root **`Sorted PDFs/`**. Subfolders are again treated as **board names**. (Useful for manually curated classification.)
- **Export (v0.4):** from the **Overview** tab click **Export…** to export selected data in **CSV**, **XLSX** *(optional, requires `openpyxl`)*, and/or **formatted TXT**. You can restrict export to specific boards and choose which fields to include.
- **Board detection by folder name**: the immediate subfolder under `PDF/` or `Sorted PDFs/` is taken as the **board** (e.g., `PDF/CaSQB/...`). The app **ignores `__archive__/`** anywhere under the roots.
- **Signature date normalization**: signature date is normalized into `YYYY-MM-DD` when possible (robust parsing of common DD/MM/YYYY, MM-DD-YYYY, etc.).

> **Note:** The app UI is English; this README is in English for consistency.

---

## Folder Structure
```
project-root/
  PDF/
    CaSQB/                # Board folder (Czech and Slovak Quality Board)
      *.pdf
    BCS/
      *.pdf
    __archive__/          # Ignored by the scanner
  Sorted PDFs/            # Optional secondary root for “Sorted PDFs” tab
    CaSQB/
      *.pdf
    ...
  app/
    main_window.py
    pdf_parser.py
    pdf_scanner.py
    istqb_boards.py
    themes.py
  main.py
  README.md
  .gitignore
```
- **`PDF/`** is the **default root** scanned by **Overview** and **PDF Browser**.  
- **`Sorted PDFs/`** is **optional**; if present, **Sorted PDFs** tab scans it similarly to `PDF/`.

---

## Parsed Fields
From the official application form the app extracts:
1. **Application Type** (`New Application` / `Additional Recognition`, radio button)
2. **Name of University, High-, or Technical School** (Institution name)
3. **Name of candidate**
4. **Wished Recognitions** (checkboxes): **Academia Recognition**, **Certified Recognition**
5. **Contact details for information exchange**:
   - Full Name
   - Email Address
   - Phone Number
   - Postal Address
6. **Declaration and Consent → Date** (signature date; normalized to `YYYY-MM-DD`)
7. **Eligibility Evidence** (shown fully in **PDF Browser**, hidden in Overview by default):
   - Description of how ISTQB syllabi are integrated in the curriculum
   - List of courses/modules
   - Proof of ISTQB certifications (if any)
   - University website links (if any)
   - Any additional relevant information or documents (if any)
8. **File name** (file basename; full path is kept internally)

> The parser is resilient to minor PDF anomalies. If a particular PDF is malformed, fields may be incomplete.

---

## Export (v0.4)
Open the **Overview** tab and click **Export…**. The dialog allows you to:
- **Formats:** select one or more of **CSV**, **XLSX**, **TXT (formatted)**.  
  - **XLSX** requires optional dependency **`openpyxl`**. If it’s not installed, CSV/TXT still export.
- **Boards:** **All** (default) or a **custom list**. Only boards that actually contain PDFs under `PDF/` (excluding `__archive__/`) are offered.
- **Fields:** choose which columns to export (all selected by default).

After confirming, a single **Save As** dialog appears. Your chosen path (e.g., `…/export.csv`) is used as a **base**; all selected formats are saved side by side (e.g., `export.csv`, `export.xlsx`, `export.txt`).

**CSV:** UTF‑8 with headers and standard separators.  
**XLSX:** single sheet **“Export”**, bold header, basic auto‑width.  
**TXT:** fixed‑width, header + separator + rows.

---

## Installation (macOS)
Tested on macOS (Apple Silicon & Intel). Recommended **Python 3.10+**.

```bash
# 1) Create and activate a virtualenv
python3 -m venv .venv
source .venv/bin/activate

# 2) Install runtime deps (minimal)
pip install --upgrade pip
pip install PySide6

# 3) (Optional) for XLSX export
pip install openpyxl

# 4) Run the app
python main.py
```

If you prefer module form:
```bash
python -m pip install -U PySide6
python -m main
```

---

## Usage
1. Place your PDFs into `./PDF/<BoardName>/…` (use actual ISTQB board names where possible; e.g., `CaSQB` for *Czech and Slovak Quality Board*).  
2. Launch the app (`python main.py`).  
3. **Overview** shows the aggregated table; **PDF Browser** shows per‑file details (including Eligibility).  
4. *(Optional)* Use **Sorted PDFs** tab if you maintain a separate `Sorted PDFs/` tree.  
5. Use **Export…** in **Overview** to produce CSV/XLSX/TXT for selected boards/fields.

### Notes
- Subfolders named `__archive__` are ignored in all scans.
- Some PDFs may contain broken objects; these are skipped gracefully and shown partially when possible.
- Board detection is **by folder name**; it does **not** validate against a registry.

---

## Troubleshooting
- **A PDF shows incomplete/incorrect values:**  
  Try opening it via **PDF Browser** to inspect all fields. If the file is malformed, you may need to correct the data externally before export.
- **XLSX not exported:**  
  Ensure **`openpyxl`** is installed (`pip install openpyxl`). CSV/TXT exports don’t need it.
- **HiDPI rendering warnings:**  
  Some Qt attributes are marked deprecated on newer Qt builds; they can be safely ignored for now.

---

## .gitignore (suggested)
```
# macOS
.DS_Store

# Python
__pycache__/
*.pyc
.venv/

# App data
PDF/
Sorted PDFs/
recognition_db.json

# Editors/IDE
.vscode/
.idea/
```

> If you prefer to **keep PDFs in the repo**, remove `PDF/` (and `Sorted PDFs/`) from `.gitignore`.

---

## Create & connect a new private GitHub repo (via `gh`)
```bash
# Initialize local repo
git init
git add .
git commit -m "feat: initial import (v0.4)"

# Create a new private GitHub repo in the current directory
gh repo create ISTQB-Academia-PDF-Aggregator --private --source=. --remote=origin --push

# Tag the current version
git tag v0.4
git push --tags
```

---

## Changelog

- **0.6a — 2025-11-12**
  - **fix:** Preserve selection in **Overview** during auto-rescan (selection no longer clears before exporting to *Sorted PDFs*).
  - **fix:** Changed Czech prompt *"Vyber alespoň jeden řádek."* to English: *"Please select at least one row."* in the export-to-sorted action.

- **0.5 (2025-11-12):**
  - **feat:** Redesigned 'Sorted PDFs' tab into a persistent database-backed manager.
  - **feat:** Added a `database.json` to store, edit, and track data for sorted PDFs.
  - **feat:** Implemented 'Export to Sorted PDFs' action from the Overview tab using multi-select.
  - **feat:** Enabled data editing and deletion for records in the Sorted PDFs view.
  - **fix:** Corrected `SyntaxError` in the database update method.
- **0.4 — 2025-11-12**
  - **Export dialog** in **Overview**: CSV/XLSX/TXT, board selection (from actual PDF folders), selectable fields, single Save As for multiple formats.
- **0.3h — 2025‑11‑11**
  - Baseline: Overview + PDF Browser, grouped columns, Yes/No icons for Wished Recognitions, signature date normalization, `__archive__` ignored.
  - Optional **Sorted PDFs** tab support (reads `Sorted PDFs/` if present).
---

## License
This project is provided “as is” for internal usage. Add your preferred license if needed.

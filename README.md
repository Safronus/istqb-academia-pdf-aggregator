# ISTQB Academia PDF Aggregator (PySide6)

**Version:** 0.3 — 2025-11-11

A macOS-oriented PySide6 GUI app (English UI) that scans a local `PDF/` folder (including subfolders) for ISTQB® Academia Recognition Program application PDFs and aggregates key fields into an overview table. A second tab provides a PDF-focused browser filtered to `.pdf` files only.

> Designed for minimal dependencies, dark-theme by default, HiDPI-friendly, and safe to run on macOS.

---

## What's new in 0.2

- **Export** currently visible rows to **CSV** and **XLSX** (menu actions).
- **Boards list updated**: *CaSTB* renamed to **CaSQB — Czech and Slovak Quality Board** and the known boards set was expanded.
- README updated; `.gitignore` added to ignore `PDF/` (incl. subfolders), `.DS_Store`, caches, and typical Python artifacts.

---

## Features

- **Automatic scanning** of `PDF/` subfolder (relative to the script location), recursively.
- **Extraction of key fields** from each PDF:
  1. *Application Type* (New Application / Additional Recognition)
  2. *Name of University, High-, or Technical School*
  3. *Name of candidate*
  4. *Recognition requested* — checkboxes (bools): Academia and Certified
  5. *Contact details*: Full Name, Email Address, Phone Number, Postal Address
  6. *Date* — field above signature (signature date)
  7. *Proof of ISTQB certifications* — free text (if present)
- **Board awareness**: Each immediate subfolder under `PDF/` is considered the ISTQB Member Board name. The app keeps an internal list of known boards and flags unknown names without blocking your workflow.
- **English UI**, **dark theme**, responsive layout for Retina/HiDPI displays.
- **Open PDF** action on any selected row/file.
- **Export CSV/XLSX** of the **currently visible** rows in the Overview table.

> The parser is tailored to the *ISTQB® Academia Recognition Program – Application Form* structure and uses robust heuristics to handle typical exports. You can extend `app/pdf_parser.py` if your forms differ.

---

## Requirements

- Python **3.10+** (tested with 3.11)
- macOS (works on other OSes but the target is macOS)
- Pip packages:
  - `PySide6`
  - `pypdf`
  - `openpyxl` *(for XLSX export)*

Install into a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install PySide6 pypdf openpyxl
```

---

## Project Structure

```
project-root/
├─ main.py
├─ app/
│  ├─ __init__.py
│  ├─ main_window.py
│  ├─ pdf_scanner.py
│  ├─ pdf_parser.py
│  ├─ istqb_boards.py
│  └─ themes.py
└─ PDF/                 # put your PDFs here; subfolders = ISTQB Member Boards
```

Create the `PDF/` folder next to `main.py` and place your PDFs inside, organized by board-specific subfolders. Example:

```
PDF/
├─ CaSQB/
│  └─ ISTQB Academia Recognition Program - Application Form v1.0_PetrZacek-SIG.pdf
└─ UKITB/
   └─ sample.pdf
```

---

## Run

From the project root:

```bash
python main.py
```

Optional: override the default `PDF` root via CLI arg:

```bash
python main.py --pdf-root /path/to/another/folder
```

---

## Usage

- **Overview tab**: aggregated table of all parsed PDFs. Use **Board** drop-down and **Search** box to narrow rows. Double-click a row or use **Open PDF** to open the file.
- **PDF Browser tab**: a left-hand directory tree filtered to `.pdf` files; selecting a file shows the parsed details on the right.
- **Export**: use the menu actions **Export CSV** or **Export XLSX** to export the **currently visible rows** in the Overview table.

---

## Notes on ISTQB Member Boards

The application ships with an initial list of well-known ISTQB Member Boards for validation hints. Board names are **not enforced**; unknown names will be allowed but shown as *Unverified*. Update `app/istqb_boards.py` to add/remove boards as needed without changing the rest of the code.  
Changed: **CaSTB ➜ CaSQB (Czech and Slovak Quality Board)**.

---

## Limitations & Tips

- PDF text extraction depends on how the PDF was produced. The app uses `pypdf` and includes heuristics for both "typed" and "scanned" variants where text exists. Image-only scans without embedded text will not parse.
- If a field is not detected, it appears as empty; you can still open the file from the UI to review it manually.
- For very large folders, first run may take longer due to parsing; subsequent rescans reuse the same logic.

---

## Development

- Code style: PEP 8, type hints, small cohesive functions.
- No risky operations, no shell calls.
- Cross-platform safe paths via `pathlib`.

---

## Versioning

- **0.2** — export CSV/XLSX; boards list updated (CaSQB).  
- 0.1 — initial release with Overview and PDF Browser tabs, recursive scan, parsing of required fields, board awareness, dark theme.

---

## Changelog

### 0.2c — 2025-11-11
- **Fix:** Read **Contact details (Section 4)** from AcroForm text fields: *Full Name*, *Email Address*, *Phone Number*, *Postal Address* (field names mapped from `Contact name`, `Contact email`, `Contact phone`, `Postal address`). Falls back to text heuristics if fields missing.


### 0.2b — 2025-11-11
- **Fix:** Read *Name of University, High-, or Technical School* and *Name of candidate* from **Section 2** form fields (AcroForm text fields), not from plain text heuristics.


### 0.2a — 2025-11-11
- **Fix:** Correctly read *Application Type* from PDF **radio buttons** (prefers AcroForm `/V` value). Prevents misclassification of "New Application" vs "Additional Recognition".
- Clarified: Board name is derived from the **immediate subfolder** under `PDF/` that contains the file.


### 0.2 — 2025-11-11
- Add export of visible rows to CSV and XLSX.
- Update ISTQB boards list (rename CaSTB → CaSQB; expand set).

### 0.1 — initial
- Overview and PDF Browser tabs.
- Recursive scan of `PDF/` root and extraction of required fields.
- Board awareness with known-board hinting.
- Dark theme by default; English UI.

---

## License

© 2025. Provided as-is. ISTQB® is a registered trademark of its respective owner.


### 0.2d — 2025-11-11
- **Fix:** Read **Signature Date** from Section 6 *Declaration and Consent* field (`Signature Date_af_date`) instead of generic "Date" heuristics.
- **Feature:** Parse Section 5 *Eligibility Evidence* fields from AcroForm text fields:
  1) *Description of how ISTQB syllabi are integrated in the curriculum* (field name: `Descriptino of how syllabi are integrated`),
  2) *List of courses/modules* (`List of courses and modules`),
  3) *Proof of ISTQB certifications (if any)* (`Proof of certifications`),
  4) *University website links (if any)* (`University website links`),
  5) *Any additional relevant information or documents (if any)* (`Additional relevant information or documents`).
- UI detail panel shows these fields.

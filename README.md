# ISTQB Academia PDF Aggregator

## Version 0.7f — 2025-11-12
### TXT export (Overview & Sorted) — unified rich report
- New **TXT report layout** shared by **Overview** and **Sorted PDFs** exports:
  - **Heading** per record (Institution — Candidate — [Board] or File name).
  - **Basic info** block (Board, Application Type, Institution, Candidate, Signature Date, File name).
  - **Sections** with **bulleted subsections** (only when data present):
    - Recognition (Academia, Certified)
    - Contact (Full Name, Email, Phone, Postal Address)
    - Curriculum (Syllabi Integration, Courses/Modules)
    - Evidence (Proof of ISTQB Certifications, Additional Info/Documents)
    - Links (University Links)
- CSV/XLSX remain unchanged.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

### Recent
- 0.7f — TXT report layout (Overview & Sorted) with headings, basic info, sections & bullets.
- 0.7e — Sorted export dialog/options mirror Overview.
- 0.7d — Sorted export columns/filters/engines same as Overview.
- 0.7c — add `ed_sigdate` alias.
- 0.7b — add `ed_rec_acad/ed_rec_cert` + aliases.
- 0.7a — fix field names/aliases in Sorted builder.
- 0.6c — scanner: strip leading '/' from Application Type.

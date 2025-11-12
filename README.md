# ISTQB Academia PDF Aggregator

## Version 0.9a — 2025-11-12
### Board Contacts tab — help & CSV template
- Tab renamed to **Board Contacts**.
- Added **Help** with an example CSV for importing contacts.
- CSV template columns (case-insensitive): `board, full_name, email`.
- You can **preview** the template in-app and **save** it via “Save CSV template…”.

### CSV example
```csv
board,full_name,email
ATB,Contact for ATB,atb-liaison@example.org
CSTB,Contact for CSTB,cstb-liaison@example.org
ISTQB,Contact for ISTQB,istqb-liaison@example.org
```

### Notes
- JSON persistence remains at `contacts.json` (git-ignored).
- Fitting of columns remains automatic.
- No behavior changes to other tabs.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```


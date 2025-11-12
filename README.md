# ISTQB Academia PDF Aggregator

## Version 0.10b — 2025-11-12
### Fix: Recognized People — candidates from Sorted DB
- **Add dialog** now discovers candidates robustněji:
  - prefers in-memory attrs: `self.sorted_db`, `_sorted_db`, `sorted_records`, `_sorted_records` (list/dict),
  - can **optionally** read common JSON paths if present (e.g. `sorted_db.json`, `sorted_records.json`, `Sorted PDFs/sorted_db.json`),
  - as a **fallback**, it can derive candidates z **Overview** tabulky (pokud není k dispozici Sorted DB).
- Přepínač **From Sorted PDFs** zůstává dostupný; pokud jsou kandidáti nalezeni, combobox je zaplněn.

### Notes
- Minimal-change: bez zásahu do stávající logiky tabů.
- Doporučení: pokud používáte vlastní strukturu/atribut Sorted DB, předejte mi prosím název/ukázku a rád doplním přesné mapování.

---

## macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

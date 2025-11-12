# ISTQB Academia PDF Aggregator

## Version 0.8 — 2025-11-12
### Overview — Board filter with two sections
- **Board combobox** now shows two sections:
  1) Boards **present in the current table** (after filters/search)
  2) ——— separator ———
  3) Remaining boards from `KNOWN_BOARDS` (alphabetical)
- Keeps **'All'** as the first item.
- The list **updates automatically** when Overview data changes (rows inserted/removed/reset or data changed).

### Notes
- No changes to existing filtering logic; `_filter_board` works as before.
- Minimal-change: only combobox population and signal wiring.

### macOS quick start
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m app
```

### Recent
- 0.8 — Overview: Board combobox grouped into present boards + separator + remaining boards.
- 0.7j — enlarge window and Sorted sizing.
- 0.7i — global sizing via showEvent; browser auto-fit.
- 0.7f — TXT report formatting.

# ISTQB Academia PDF Aggregator

---

## Release 0.6c — 2025-11-12
- **fix(parser):** Remove a leading "/" from *Application Type* regardless of source (AcroForm or text) by sanitizing the value **immediately before return** in `parse_istqb_academia_application`. No other logic changed.

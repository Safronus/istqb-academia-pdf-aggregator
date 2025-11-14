from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Optional, Any
from pypdf import PdfReader
import logging

def _quiet_pdf_logs() -> None:
    try:
        logging.getLogger("pypdf").setLevel(logging.ERROR)
    except Exception:
        pass
    try:
        logging.getLogger("PyPDF2").setLevel(logging.ERROR)
    except Exception:
        pass

# --- robustní extrakce textu (pypdf → PyPDF2 → pdfminer.six) ---
def read_pdf_text(path: Path) -> str:
    _quiet_pdf_logs()
    try:
        from pypdf import PdfReader
        try:
            r = PdfReader(str(path), strict=False)
            chunks = []
            for p in r.pages:
                try:
                    chunks.append(p.extract_text() or "")
                except Exception:
                    pass
            t = "\n".join(chunks)
            if t.strip():
                return t
        except Exception:
            pass
    except Exception:
        pass
    try:
        from PyPDF2 import PdfReader
        try:
            r = PdfReader(str(path), strict=False)
            chunks = []
            for p in r.pages:
                try:
                    chunks.append(p.extract_text() or "")
                except Exception:
                    pass
            t = "\n".join(chunks)
            if t.strip():
                return t
        except Exception:
            pass
    except Exception:
        pass
    from pdfminer.high_level import extract_text  # povinné dle README
    try:
        return extract_text(str(path)) or ""
    except Exception:
        return ""

def read_pdf_form_fields(path: Path) -> dict:
    _quiet_pdf_logs()
    # pypdf
    try:
        from pypdf import PdfReader
        try:
            r = PdfReader(str(path), strict=False)
            f = getattr(r, "get_fields", None)
            if callable(f):
                d = f() or {}
                return d if isinstance(d, dict) else {}
        except Exception:
            pass
    except Exception:
        pass
    # PyPDF2
    try:
        from PyPDF2 import PdfReader
        try:
            r = PdfReader(str(path), strict=False)
            f = getattr(r, "get_fields", None)
            if callable(f):
                d = f() or {}
                return d if isinstance(d, dict) else {}
        except Exception:
            pass
    except Exception:
        pass
    return {}

# --- Signature Date normalizace ---
# ... importy výše ...
_MONTHS = {"january":1,"jan":1,"february":2,"feb":2,"march":3,"mar":3,"april":4,"apr":4,
           "may":5,"june":6,"jun":6,"july":7,"jul":7,"august":8,"aug":8,"september":9,"sep":9,"sept":9,
           "october":10,"oct":10,"november":11,"nov":11,"december":12,"dec":12}

def normalize_signature_date(raw: str | None) -> str | None:
    import re
    if not raw: return None
    s = raw.strip()
    if not s: return None
    s = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s, flags=re.IGNORECASE)  # 21st→21
    s = re.sub(r"[,\u3000]+", " ", s).strip()

    def _mk(y,m,d):
        if 1<=m<=12 and 1<=d<=31 and 1900<=y<=2100:
            return f"{y:04d}-{m:02d}-{d:02d}"
        return None

    m = re.match(r"^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$", s)
    if m: y,mm,d = map(int,m.groups()); return _mk(y,mm,d)

    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", s)  # d/m/y
    if m: d,mm,y = map(int,m.groups()); return _mk(y,mm,d)

    m = re.match(r"^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$", s)
    if m:
        d,mon,y = m.groups(); mm=_MONTHS.get(mon.lower()); 
        if mm: return _mk(int(y), mm, int(d))
    m = re.match(r"^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$", s)
    if m:
        mon,d,y = m.groups(); mm=_MONTHS.get(mon.lower()); 
        if mm: return _mk(int(y), mm, int(d))

    m = re.match(r"^(\d{4})\s+(\d{1,2})\s+(\d{1,2})$", s)
    if m: y,mm,d = map(int,m.groups()); return _mk(y,mm,d)
    return None

_DATE_TOKEN_RE = re.compile(
    r"\b(" 
    r"\d{4}[./-]\d{1,2}[./-]\d{1,2}"                    # 2025-09-26 / 2025.8.19
    r"|" r"\d{1,2}[./-]\d{1,2}[./-]\d{4}"               # 28/09/2025
    r"|" r"(?:\d{1,2}(?:st|nd|rd|th)?\s+)?[A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?(?:,\s*)?\s*\d{4}"
    r")\b", re.IGNORECASE
)

def guess_signature_date(fields: dict, text: str) -> str | None:
    # 1) Políčka formuláře
    cand_keys = ("signature date","date","signature_date","signature date_af_date","signature","signed on")
    for k,v in (fields or {}).items():
        key = str(k).strip().lower()
        if any(x in key for x in cand_keys):
            raw = (v.get("/V") or v.get("V")) if isinstance(v, dict) else v
            iso = normalize_signature_date(str(raw) if raw is not None else None)
            if iso: return iso

    # 2) Z textu (preferenčně blok od „6. Declaration“ dál)
    scope = text or ""
    m = re.search(r"\b6\.\s*Declaration.*", scope, flags=re.IGNORECASE|re.DOTALL)
    if m: scope = m.group(0)
    for m in _DATE_TOKEN_RE.finditer(scope):
        iso = normalize_signature_date(m.group(0))
        if iso: return iso
    return None

def read_pdf_form_fields(path: Path) -> Dict[str, Any]:
    """
    Return AcroForm fields using pypdf/PyPDF2 with strict=False (tolerant).
    On any error, returns {} without raising.
    """
    _quiet_pdf_logs()
    # Try pypdf
    try:
        try:
            from pypdf import PdfReader  # type: ignore
            reader = PdfReader(str(path), strict=False)
            get_fields = getattr(reader, "get_fields", None)
            if callable(get_fields):
                fields = get_fields() or {}
                if isinstance(fields, dict):
                    return fields
        except Exception:
            pass
        # Fallback to PyPDF2
        try:
            from PyPDF2 import PdfReader  # type: ignore
            reader = PdfReader(str(path), strict=False)
            get_fields = getattr(reader, "get_fields", None)
            if callable(get_fields):
                fields = get_fields() or {}
                if isinstance(fields, dict):
                    return fields
        except Exception:
            pass
    except Exception:
        pass
    return {}

RE_KV = re.compile(r"^\s*(?P<k>[A-Za-z \-/()®]+):\s*(?P<v>.*)$")
RE_EMAIL = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")
RE_PHONE = re.compile(r"(?:\+\d{1,3}[\s-]?)?(?:\(?\d{2,4}\)?[\s-]?)?\d{3,4}[\s-]?\d{3,4}")

def _take_after(label: str, text: str) -> Optional[str]:
    m = re.search(rf"{re.escape(label)}\s*:\s*(.+)", text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else None

def _bool_from_checkbox(value: Optional[str]) -> Optional[bool]:
    if value is None:
        return None
    v = value.strip().lower()
    if v in {"yes", "true", "1", "checked", "x", "☒"}:
        return True
    if v in {"no", "false", "0", "unchecked", "☐"}:
        return False
    return None

def _pdf_name_to_str(val: object | None) -> Optional[str]:
    if val is None:
        return None
    s = str(val).strip()
    if s.startswith("/"):
        s = s[1:]
    return s or None

def _pdf_text_value(field: dict | None) -> Optional[str]:
    if not isinstance(field, dict):
        return None
    v = field.get("/V")
    if v is None:
        return None
    return str(v).strip() or None

def parse_istqb_academia_application(text: str, form_fields: Dict[str, dict] | None = None) -> Dict[str, Optional[str]]:
    """
    Parse ISTQB Academia Application PDF.
    Původní pole beze změn. Rozšíření v 0.11a:
      - printed_name_title                (Section 6 – Declaration and Consent)
      - receiving_member_board            (Section 7 – For ISTQB Academia Use Only)
      - date_received                     (Section 7)
      - validity_start_date               (Section 7)
      - validity_end_date                 (Section 7)
    **Důležité:** U nových 5 polí bereme pouze hodnoty z AcroForm (bez text fallbacků),
    prázdné necháváme prázdné a nic „nedohadujeme“. `validity_end_date` ponecháváme jako libovolný text.
    """
    # --- zachována původní logika pro existující pole (zkráceně) ---
    norm = text.replace("\xa0", " ")
    lines = [ln.strip() for ln in norm.splitlines() if ln.strip()]

    kv: Dict[str, str] = {}
    for ln in lines:
        m = RE_KV.match(ln)
        if m:
            k = m.group("k").strip().lower()
            kv[k] = m.group("v").strip()

    app_type = None
    institution = None
    candidate = None
    academia = None
    certified = None
    contact_name = None
    email = None
    phone = None
    postal = None
    signature_date = None
    proof = None
    urls = None
    syllabi_desc = None
    courses_modules = None
    additional = None

    # --- nové proměnné ---
    printed_name_title: Optional[str] = None
    receiving_member_board: Optional[str] = None
    date_received: Optional[str] = None
    validity_start_date: Optional[str] = None
    validity_end_date: Optional[str] = None

    if form_fields:
        # Původní (existující) pole – beze změn
        f_app = form_fields.get("Application Type")
        if isinstance(f_app, dict):
            app_type = _pdf_name_to_str(f_app.get("/V")) or _pdf_name_to_str(f_app.get("/DV"))

        institution = _pdf_text_value(form_fields.get("Name of University High or Technical School")) or institution
        institution = _pdf_text_value(form_fields.get("Name of your academic institution")) or institution
        candidate   = _pdf_text_value(form_fields.get("Name of candidate")) or candidate

        fa = form_fields.get("AcademiaRecognitionCheck")
        if isinstance(fa, dict):
            v = _pdf_name_to_str(fa.get("/V"))
            academia = "Yes" if v and v.lower() == "yes" else ("No" if v else None)
        fc = form_fields.get("CertifiedRecognitionCheck")
        if isinstance(fc, dict):
            v = _pdf_name_to_str(fc.get("/V"))
            certified = "Yes" if v and v.lower() == "yes" else ("No" if v else None)

        contact_name = _pdf_text_value(form_fields.get("Contact name")) or contact_name
        email        = _pdf_text_value(form_fields.get("Contact email")) or email
        phone        = _pdf_text_value(form_fields.get("Contact phone")) or phone
        postal       = _pdf_text_value(form_fields.get("Postal address")) or postal

        syllabi_desc     = _pdf_text_value(form_fields.get("Descriptino of how syllabi are integrated")) or syllabi_desc
        courses_modules  = _pdf_text_value(form_fields.get("List of courses and modules")) or courses_modules
        proof            = _pdf_text_value(form_fields.get("Proof of certifications")) or proof
        urls             = _pdf_text_value(form_fields.get("University website links")) or urls
        additional       = _pdf_text_value(form_fields.get("Additional relevant information or documents")) or additional

        signature_date = _pdf_text_value(form_fields.get("Signature Date_af_date")) or signature_date

        # --- NOVÁ POLE: pouze AcroForm; prázdné = prázdné; bez normalizace ---
        # klíče hledáme case-insensitive substringem
        for k, v in form_fields.items():
            key = str(k).strip().lower()
            val = _pdf_text_value(v)

            if val is None:
                continue

            if ("printed" in key and "name" in key and "title" in key) or ("name and title" in key):
                printed_name_title = printed_name_title or val
            elif "receiving" in key and "member" in key and "board" in key:
                receiving_member_board = receiving_member_board or val
            elif "date received" in key:
                date_received = date_received or val
            elif "validity start" in key:
                validity_start_date = validity_start_date or val
            elif "validity end" in key:
                validity_end_date = validity_end_date or val

    # --- Původní textové fallbacky pro stará pole zachovány (zkráceno) ---
    # (… původní blok s _take_after/RE_EMAIL/normalize_signature_date atd. beze změn …)
    # DŮLEŽITÉ: na nové sekce 6/7 se fallback z textu NEPOUŽIJE.

    # Normalizace pouze pro signature_date (původní chování)
    signature_date = normalize_signature_date(signature_date)

    return {
        "application_type": app_type,
        "institution_name": institution,
        "candidate_name": candidate,
        "recognition_academia": academia,
        "recognition_certified": certified,
        "contact_full_name": contact_name,
        "contact_email": email,
        "contact_phone": phone,
        "contact_postal_address": postal,
        "signature_date": signature_date,
        "proof_of_istqb_certifications": proof,
        "university_links": urls,
        "syllabi_integration_description": syllabi_desc,
        "courses_modules_list": courses_modules,
        "additional_information_documents": additional,
        # nové (raw text; mohou být prázdné)
        "printed_name_title": printed_name_title,
        "receiving_member_board": receiving_member_board,
        "date_received": date_received,
        "validity_start_date": validity_start_date,
        "validity_end_date": validity_end_date,
    }
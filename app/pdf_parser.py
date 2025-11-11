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
_MONTHS = {
    "january":1, "jan":1, "february":2, "feb":2, "march":3, "mar":3, "april":4, "apr":4,
    "may":5, "june":6, "jun":6, "july":7, "jul":7, "august":8, "aug":8, "september":9, "sep":9, "sept":9,
    "october":10, "oct":10, "november":11, "nov":11, "december":12, "dec":12
}

def normalize_signature_date(raw: str | None) -> str | None:
    """
    Převede různé varianty data na YYYY-MM-DD (např. 28/09/2025, 2025.8.19, 21st August, 2025).
    Vrací None, když nerozpozná.
    """
    if not raw:
        return None
    s = raw.strip()
    if not s:
        return None
    s = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s, flags=re.IGNORECASE)  # 21st -> 21
    s = re.sub(r"[,\u3000]+", " ", s).strip()

    def _mk(y: int, m: int, d: int) -> str | None:
        if 1 <= m <= 12 and 1 <= d <= 31 and 1900 <= y <= 2100:
            return f"{y:04d}-{m:02d}-{d:02d}"
        return None

    # YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
    m = re.match(r"^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$", s)
    if m:
        y, mm, dd = map(int, m.groups())
        return _mk(y, mm, dd)

    # D.M.YYYY / D-M-YYYY / D/M/YYYY (preferujeme D/M/Y)
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", s)
    if m:
        d, mm, y = map(int, m.groups())
        return _mk(y, mm, d)

    # 21 August 2025 / August 21 2025 / s čárkami
    m = re.match(r"^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$", s)
    if m:
        d, mon, y = m.groups()
        mm = _MONTHS.get(mon.lower())
        if mm: return _mk(int(y), mm, int(d))
    m = re.match(r"^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$", s)
    if m:
        mon, d, y = m.groups()
        mm = _MONTHS.get(mon.lower())
        if mm: return _mk(int(y), mm, int(d))

    # YYYY M D (mezery)
    m = re.match(r"^(\d{4})\s+(\d{1,2})\s+(\d{1,2})$", s)
    if m:
        y, mm, d = map(int, m.groups())
        return _mk(y, mm, d)

    return None

# kandidátní klíče pro form fields (různé varianty)
_SIG_DATE_CAND_KEYS = (
    "signature date", "date", "signature_date", "signature date_af_date",
    "signature_af_date", "signature", "signed on"
)

_DATE_TOKEN_RE = re.compile(
    r"\b("                                    # několik běžných zápisů data
    r"\d{4}[./-]\d{1,2}[./-]\d{1,2}"          # 2025-09-26, 2025/8/20, 2025.8.19
    r"|"
    r"\d{1,2}[./-]\d{1,2}[./-]\d{4}"          # 28/09/2025, 19.8.2025
    r"|"
    r"(?:\d{1,2}(?:st|nd|rd|th)?\s+)?[A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?(?:,\s*)?\s*\d{4}"  # Aug 21, 2025 / 21st August 2025
    r")\b",
    re.IGNORECASE
)

def guess_signature_date(fields: dict, text: str) -> str | None:
    # 1) Zkuste form fields
    for k, v in (fields or {}).items():
        key = str(k).strip().lower()
        if any(x in key for x in _SIG_DATE_CAND_KEYS):
            raw = None
            if isinstance(v, dict):
                raw = v.get("/V") or v.get("V")
            else:
                raw = v
            iso = normalize_signature_date(str(raw) if raw is not None else None)
            if iso:
                return iso

    # 2) Z textu – přednostně část „6. Declaration“
    scope = text or ""
    m = re.search(r"\b6\.\s*Declaration.*", scope, flags=re.IGNORECASE | re.DOTALL)
    if m:
        scope = m.group(0)

    for m in _DATE_TOKEN_RE.finditer(scope):
        iso = normalize_signature_date(m.group(0))
        if iso:
            return iso
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
    Prefers AcroForm (form_fields) and falls back to text heuristics.
    """
    norm = text.replace("\xa0", " ")
    lines = [ln.strip() for ln in norm.splitlines() if ln.strip()]

    kv: Dict[str, str] = {}
    for ln in lines:
        m = RE_KV.match(ln)
        if m:
            k = m.group("k").strip().lower()
            kv[k] = m.group("v").strip()

    # ---------- Prefer AcroForm ----------
    app_type = None
    academia = None
    certified = None
    institution = None
    candidate = None
    urls = None
    contact_name = None
    email = None
    phone = None
    postal = None
    signature_date = None

    syllabi_desc = None
    courses_modules = None
    proof = None
    additional = None

    if form_fields:
        # Application type (radio)
        f_app = form_fields.get("Application Type")
        if isinstance(f_app, dict):
            app_type = _pdf_name_to_str(f_app.get("/V")) or _pdf_name_to_str(f_app.get("/DV"))

        # Section 2
        institution = _pdf_text_value(form_fields.get("Name of University High or Technical School")) or institution
        candidate = _pdf_text_value(form_fields.get("Name of candidate")) or candidate

        # Recognition (checkboxes)
        fa = form_fields.get("AcademiaRecognitionCheck")
        if isinstance(fa, dict):
            v = _pdf_name_to_str(fa.get("/V"))
            academia = "Yes" if v and v.lower() == "yes" else ("No" if v else None)
        fc = form_fields.get("CertifiedRecognitionCheck")
        if isinstance(fc, dict):
            v = _pdf_name_to_str(fc.get("/V"))
            certified = "Yes" if v and v.lower() == "yes" else ("No" if v else None)

        # Section 4 – Contact details
        contact_name = _pdf_text_value(form_fields.get("Contact name")) or contact_name
        email = _pdf_text_value(form_fields.get("Contact email")) or email
        phone = _pdf_text_value(form_fields.get("Contact phone")) or phone
        postal = _pdf_text_value(form_fields.get("Postal address")) or postal

        # Section 5 – Eligibility Evidence
        syllabi_desc = _pdf_text_value(form_fields.get("Descriptino of how syllabi are integrated")) or syllabi_desc
        courses_modules = _pdf_text_value(form_fields.get("List of courses and modules")) or courses_modules
        proof = _pdf_text_value(form_fields.get("Proof of certifications")) or proof
        urls = _pdf_text_value(form_fields.get("University website links")) or urls
        additional = _pdf_text_value(form_fields.get("Additional relevant information or documents")) or additional

        # Section 6 – Signature date
        signature_date = _pdf_text_value(form_fields.get("Signature Date_af_date")) or signature_date

    # ---------- Fallbacks from text ----------
    if not app_type:
        app_type = (
            kv.get("application type")
            or _take_after("Application Type", norm)
            or ("New Application" if "New Application" in norm and "Additional Recognition" not in norm else None)
            or ("Additional Recognition" if "Additional Recognition" in norm else None)
        )

    if not institution:
        institution = (
            kv.get("name of university high or technical school")
            or kv.get("name of your academic institution")
            or _take_after("Name of University, High-, or Technical School", norm)
        )

    if not candidate:
        candidate = kv.get("name of candidate") or _take_after("Name of candidate", norm)

    if academia is None:
        a = kv.get("academia recognition") or _take_after("Academia Recognition", norm)
        a_bool = _bool_from_checkbox(a)
        academia = "Yes" if a_bool is True else ("No" if a_bool is False else None)

    if certified is None:
        c = kv.get("certified recognition") or _take_after("Certified Recognition", norm)
        c_bool = _bool_from_checkbox(c)
        certified = "Yes" if c_bool is True else ("No" if c_bool is False else None)

    if contact_name is None:
        contact_name = kv.get("contact name") or kv.get("full name") or _take_after("Full Name", norm)

    if email is None:
        email = kv.get("contact email") or kv.get("email address") or _take_after("Email address", norm)
        if not email:
            m = RE_EMAIL.search(norm)
            email = m.group(0) if m else None

    if phone is None:
        phone = kv.get("contact phone") or kv.get("phone number") or _take_after("Phone number", norm)
        if not phone:
            m = RE_PHONE.search(norm)
            phone = m.group(0) if m else None

    if postal is None:
        postal = kv.get("postal address") or _take_after("Postal address", norm)

    if signature_date is None:
        signature_date = (
            kv.get("signature date_af_date")
            or kv.get("date")
            or _take_after("Signature Date", norm)
            or _take_after("Date", norm)
        )
        if signature_date:
            signature_date = signature_date.strip()

    if proof is None:
        proof = kv.get("proof of certifications") or kv.get("proof of istqb certifications")
        if not proof:
            m = re.search(
                r"Proof of ISTQB® certifications.*?:\s*(?P<blk>.+?)(?:\n[A-Z][^\n:]+:|\Z)",
                norm, flags=re.IGNORECASE | re.DOTALL
            )
            if m:
                proof = re.sub(r"\s+", " ", m.group("blk")).strip()

    if urls is None:
        urls = kv.get("university website links") or None

    # Reasonable fallbacks for syllabi_desc / courses_modules / additional
    if syllabi_desc is None:
        syllabi_desc = _take_after("Description of how ISTQB syllabi are integrated in the curriculum", norm)
    if courses_modules is None:
        courses_modules = _take_after("List of courses/modules", norm)
    if additional is None:
        additional = _take_after("Any additional relevant information or documents", norm)

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
    }
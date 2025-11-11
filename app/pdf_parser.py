from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Optional, Any
from pypdf import PdfReader


def read_pdf_form_fields(path: Path) -> Dict[str, Any]:
    """
    Return AcroForm fields as a dict using pypdf/PyPDF2 if available.
    Keys are field names; values are raw field dicts (PyPDF2/pypdf model).
    If no form or error, returns {}.
    """
    try:
        try:
            from pypdf import PdfReader  # type: ignore
        except Exception:
            from PyPDF2 import PdfReader  # type: ignore
        reader = PdfReader(str(path))
        get_fields = getattr(reader, "get_fields", None)
        if callable(get_fields):
            fields = get_fields() or {}
            return fields
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
    """
    Convert a PDF Name object like '/New Application' or '/Yes' to plain text.
    Works with both pypdf and PyPDF2, tolerates plain strings as well.
    """
    if val is None:
        return None
    s = str(val).strip()
    if s.startswith("/"):
        s = s[1:]
    return s or None

def parse_istqb_academia_application(text: str, form_fields: Dict[str, dict] | None = None) -> Dict[str, Optional[str]]:
    """
    Parse target fields from ISTQB Academia Recognition Program Application PDF text.
    Prefers values from AcroForm (if provided via form_fields), falls back to regex/text heuristics.
    """
    norm = text.replace("\xa0", " ")
    lines = [ln.strip() for ln in norm.splitlines() if ln.strip()]

    # Build key-value map for "Key: Value" lines
    kv: Dict[str, str] = {}
    for ln in lines:
        m = RE_KV.match(ln)
        if m:
            k = m.group("k").strip().lower()
            kv[k] = m.group("v").strip()

    # ---------- Prefer AcroForm values where available ----------
    app_type = None
    academia = None
    certified = None

    if form_fields:
        # 1) Application Type as radio (/Btn) with value '/New Application' or '/Additional Recognition'
        f = form_fields.get("Application Type")
        if isinstance(f, dict):
            app_type = _pdf_name_to_str(f.get("/V")) or _pdf_name_to_str(f.get("/DV"))

        # 4) Recognition checkboxes
        fa = form_fields.get("AcademiaRecognitionCheck")
        if isinstance(fa, dict):
            academia = _pdf_name_to_str(fa.get("/V"))
            academia = "Yes" if academia and academia.lower() == "yes" else ("No" if academia else None)

        fc = form_fields.get("CertifiedRecognitionCheck")
        if isinstance(fc, dict):
            certified = _pdf_name_to_str(fc.get("/V"))
            certified = "Yes" if certified and certified.lower() == "yes" else ("No" if certified else None)

    # ---------- Fallbacks from text (kept pro kompatibilitu) ----------
    if not app_type:
        app_type = (
            kv.get("application type")
            or _take_after("Application Type", norm)
            or ("New Application" if "New Application" in norm and "Additional Recognition" not in norm else None)
            or ("Additional Recognition" if "Additional Recognition" in norm else None)
        )

    institution = (
        kv.get("name of university high or technical school")
        or kv.get("name of your academic institution")
        or _take_after("Name of University, High-, or Technical School", norm)
    )

    candidate = kv.get("name of candidate") or _take_after("Name of candidate", norm)

    if academia is None:
        academia_raw = (
            kv.get("academia recognition")
            or _take_after("Academia Recognition", norm)
        )
        a_bool = _bool_from_checkbox(academia_raw)
        academia = "Yes" if a_bool is True else ("No" if a_bool is False else None)

    if certified is None:
        certified_raw = (
            kv.get("certified recognition")
            or _take_after("Certified Recognition", norm)
        )
        c_bool = _bool_from_checkbox(certified_raw)
        certified = "Yes" if c_bool is True else ("No" if c_bool is False else None)

    contact_name = (
        kv.get("contact name")
        or kv.get("full name")
        or _take_after("Full Name", norm)
    )
    email = (
        kv.get("contact email")
        or kv.get("email address")
        or _take_after("Email address", norm)
    )
    if not email:
        m = RE_EMAIL.search(norm)
        email = m.group(0) if m else None

    phone = (
        kv.get("contact phone")
        or kv.get("phone number")
        or _take_after("Phone number", norm)
    )
    if not phone:
        m = RE_PHONE.search(norm)
        phone = m.group(0) if m else None

    postal = (
        kv.get("postal address")
        or _take_after("Postal address", norm)
    )

    signature_date = (
        kv.get("signature date_af_date")
        or kv.get("date")
        or _take_after("Signature Date", norm)
        or _take_after("Date", norm)
    )
    if signature_date:
        signature_date = signature_date.strip()

    proof = (
        kv.get("proof of certifications")
        or kv.get("proof of istqb certifications")
    )
    if not proof:
        m = re.search(
            r"Proof of ISTQB® certifications.*?:\s*(?P<blk>.+?)(?:\n[A-Z][^\n:]+:|\Z)",
            norm, flags=re.IGNORECASE | re.DOTALL
        )
        if m:
            proof = re.sub(r"\s+", " ", m.group("blk")).strip()

    urls = kv.get("university website links") or None

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
    }
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Optional
from pypdf import PdfReader


def read_pdf_text(path: Path) -> str:
    """Read text from a PDF using pypdf. Returns empty string on failure."""
    try:
        reader = PdfReader(str(path))
        chunks = []
        for page in reader.pages:
            try:
                chunks.append(page.extract_text() or "")
            except Exception:
                pass
        return "\n".join(chunks)
    except Exception:
        return ""


# Precompiled regex patterns to increase performance and robustness
RE_KV = re.compile(r"^\s*(?P<k>[A-Za-z \-/()®]+):\s*(?P<v>.*)$")
RE_EMAIL = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")
RE_PHONE = re.compile(r"(?:\+\d{1,3}[\s-]?)?(?:\(?\d{2,4}\)?[\s-]?)?\d{3,4}[\s-]?\d{3,4}")

def _take_after(label: str, text: str) -> Optional[str]:
    """Finds a line starting with 'label' and returns its value after colon."""
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


def parse_istqb_academia_application(text: str) -> Dict[str, Optional[str]]:
    """
    Parse target fields from ISTQB Academia Recognition Program Application PDF text.
    Returns a dict with typed/simple values where possible.
    """
    # Normalize multiple spaces and non-breaking spaces
    norm = text.replace("\xa0", " ")
    lines = [ln.strip() for ln in norm.splitlines() if ln.strip()]

    # Build a quick key->value map for common "Key: Value" pairs seen in exports
    kv: Dict[str, str] = {}
    for ln in lines:
        m = RE_KV.match(ln)
        if m:
            k = m.group("k").strip().lower()
            kv[k] = m.group("v").strip()

    # 1) Application Type
    app_type = (
        kv.get("application type")
        or _take_after("Application Type", norm)
        or ("New Application" if "New Application" in norm and "Additional Recognition" not in norm else None)
        or ("Additional Recognition" if "Additional Recognition" in norm else None)
    )

    # 2) Institution name
    institution = (
        kv.get("name of university high or technical school")
        or kv.get("name of your academic institution")
        or _take_after("Name of University, High-, or Technical School", norm)
    )

    # 3) Candidate name
    candidate = kv.get("name of candidate") or _take_after("Name of candidate", norm)

    # 4) Recognition checkboxes (Academia / Certified)
    academia_raw = (
        kv.get("academiarecognitioncheck")
        or kv.get("academia recognition")
        or _take_after("Academia Recognition", norm)
    )
    certified_raw = (
        kv.get("certifiedrecognitioncheck")
        or kv.get("certified recognition")
        or _take_after("Certified Recognition", norm)
    )
    academia = _bool_from_checkbox(academia_raw)
    certified = _bool_from_checkbox(certified_raw)

    # 5) Contact details
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

    # 6) Date (signature area)
    signature_date = (
        kv.get("signature date_af_date")
        or kv.get("date")
        or _take_after("Signature Date", norm)
        or _take_after("Date", norm)  # fall-back; may capture other dates
    )
    if signature_date:
        signature_date = signature_date.strip()

    # 7) Proof of ISTQB certifications (free text, may contain multiple lines)
    proof = (
        kv.get("proof of certifications")
        or kv.get("proof of istqb certifications")
    )
    if not proof:
        # Fallback: capture block after heading till next blank or heading
        m = re.search(
            r"Proof of ISTQB® certifications.*?:\s*(?P<blk>.+?)(?:\n[A-Z][^\n:]+:|\Z)",
            norm, flags=re.IGNORECASE | re.DOTALL
        )
        if m:
            proof = re.sub(r"\s+", " ", m.group("blk")).strip()

    # Additional: URLs block (optional)
    urls = kv.get("university website links") or None

    return {
        "application_type": app_type,
        "institution_name": institution,
        "candidate_name": candidate,
        "recognition_academia": "Yes" if academia is True else ("No" if academia is False else None),
        "recognition_certified": "Yes" if certified is True else ("No" if certified is False else None),
        "contact_full_name": contact_name,
        "contact_email": email,
        "contact_phone": phone,
        "contact_postal_address": postal,
        "signature_date": signature_date,
        "proof_of_istqb_certifications": proof,
        "university_links": urls,
    }
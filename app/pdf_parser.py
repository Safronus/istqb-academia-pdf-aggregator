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

# --- Fallback pro flattened PDF bez AcroForm polí ---
# Hodnoty jsou jen v textové vrstvě (buď jako bloky za labely, nebo zřetězené).
_RE_EMAIL_NF = re.compile(r"[\w.\-]+@[\w.\-]+\.\w+")
_RE_PHONE_NF = re.compile(r"\+?\d[\d\s().\-]{7,}\d")
_TEMPLATE_EMAIL_LOCALPARTS = {"academia.chair"}

def _find_signature_date_in_lines(lines: list[str]) -> tuple[Optional[str], Optional[int]]:
    """Vrať (ISO datum, index řádku) posledního normalizovatelného data."""
    best: tuple[Optional[str], Optional[int]] = (None, None)
    for i, ln in enumerate(lines):
        if not ln:
            continue
        iso = normalize_signature_date(ln)
        if not iso:
            m = re.search(
                r"\d{1,2}\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{1,2},?\s*\d{4}|\d{1,4}[./-]\d{1,2}[./-]\d{1,4}",
                ln,
            )
            if m:
                iso = normalize_signature_date(m.group(0))
        if iso:
            best = (iso, i)
    return best

def _extract_no_acroform_text(norm: str) -> Dict[str, Optional[str]]:
    """Konzervativní extrakce z textu pro PDF bez AcroForm polí.

    Spouští se jen když ve formuláři nejsou žádná pole; doplní pouze pole,
    která má volající prázdná. Pro naskenovaná PDF bez textové vrstvy vrací
    samé None (žádné falešné hodnoty).
    """
    out: Dict[str, Optional[str]] = {
        "institution_name": None,
        "candidate_name": None,
        "contact_full_name": None,
        "contact_email": None,
        "contact_phone": None,
        "contact_postal_address": None,
        "signature_date": None,
        "printed_name_title": None,
    }
    if not norm or not norm.strip():
        return out
    lines = [ln.strip() for ln in norm.splitlines()]

    # 1) Datum podpisu + Printed Name (řádek před datem)
    sig, di = _find_signature_date_in_lines(lines)
    out["signature_date"] = sig
    if di is not None:
        for j in range(di - 1, -1, -1):
            if lines[j] and "@" not in lines[j] and not _RE_PHONE_NF.search(lines[j]):
                out["printed_name_title"] = lines[j]
                break
    name_tokens = []
    if out["printed_name_title"]:
        name_tokens = [t for t in re.split(r"[\s/,]+", out["printed_name_title"]) if t]

    # 2) E-mail žadatele (přeskoč šablonové adresy)
    app_email = None
    for m in _RE_EMAIL_NF.finditer(norm):
        if m.group(0).split("@", 1)[0].lower() in _TEMPLATE_EMAIL_LOCALPARTS:
            continue
        app_email = m
        break

    if app_email:
        s, e, val = app_email.start(), app_email.end(), app_email.group(0)
        at_line_start = (s == 0 or norm[s - 1] in "\n\r")
        local, _, domain = val.partition("@")
        # Pokud je e-mail zřetězený se jménem (není na začátku řádku), odstraň
        # prefix odpovídající známým částem jména.
        if not at_line_start and name_tokens:
            for tok in name_tokens:
                if len(tok) > 1 and local.lower().startswith(tok.lower()) and len(local) > len(tok):
                    local = local[len(tok):].lstrip("._- ")
                    break
        out["contact_email"] = f"{local}@{domain}"

        # Bloková layout (hodnoty pod labely): jména na řádcích před e-mailem.
        if at_line_start:
            bl = [ln for ln in norm[:s].splitlines() if ln.strip()]
            if bl:
                out["contact_full_name"] = bl[-1].strip()
                if len(bl) >= 2:
                    out["candidate_name"] = bl[-2].strip()
                if len(bl) >= 3:
                    out["institution_name"] = bl[-3].strip()

        after = norm[e:]
        pm = _RE_PHONE_NF.search(after)
        if pm:
            out["contact_phone"] = pm.group(0).strip()
            rest = after[pm.end():].split("\n", 1)[0].strip()
            if rest:
                out["contact_postal_address"] = rest
            else:
                for ln in after[pm.end():].splitlines():
                    if ln.strip():
                        out["contact_postal_address"] = ln.strip()
                        break

    # 3) Doplň jméno z Printed Name, pokud chybí
    if out["printed_name_title"]:
        nm = re.split(r"[/,]", out["printed_name_title"])[0].strip()
        if not out["candidate_name"]:
            out["candidate_name"] = nm
        if not out["contact_full_name"]:
            out["contact_full_name"] = nm
    return out

def parse_istqb_academia_application(text: str, form_fields: Dict[str, dict] | None = None) -> Dict[str, Optional[str]]:
    """
    Parse ISTQB Academia Application PDF.

    Minimal-change rozšíření:
      - zachovává prioritu AcroForm polí
      - doplňuje konzervativní fallbacky z textu PDF jen pokud jsou daná pole prázdná:
        * university_links                      (sekce 5; URL deduplikované)
        * additional_information_documents      (sekce 5; text za labelem)
        * printed_name_title                    (sekce 6; pouze když obsahuje jméno kandidáta)

    Návratová struktura zůstává beze změny.
    """
    # Bezpečné normalizace
    norm = (text or "").replace("\xa0", " ")
    lines = [ln.strip() for ln in norm.splitlines() if ln.strip()]

    # Pomocné převody – spolehni se na existující utility v modulu, pokud jsou k dispozici
    def _pdf_text_value_local(field: dict | None) -> Optional[str]:
        try:
            if not isinstance(field, dict):
                return None
            v = field.get("/V") or field.get("V")
            if v is None:
                return None
            s = str(v).strip()
            return s or None
        except Exception:
            return None

    # Preferenčně čteme z AcroForm polí (beze změny)
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

    printed_name_title: Optional[str] = None
    receiving_member_board: Optional[str] = None
    date_received: Optional[str] = None
    validity_start_date: Optional[str] = None
    validity_end_date: Optional[str] = None

    if form_fields:
        def fval(*keys: str) -> Optional[str]:
            for k in form_fields.keys():
                kk = str(k).strip().lower()
                for probe in keys:
                    if probe in kk:
                        val = _pdf_text_value_local(form_fields[k])
                        if val and val.startswith("/"):
                            val = val[1:].lstrip()
                        if val:
                            return val
            return None

        app_type = fval("application type")
        institution = fval("name of university", "name of your academic institution", "high", "technical school")
        candidate = fval("name of candidate")

        # checkboxy – vracíme "Yes"/"No"
        def checkbox_truthy(raw: Optional[str]) -> Optional[bool]:
            if raw is None:
                return None
            s = raw.strip().lower()
            if s in {"yes", "on", "true", "1", "/yes"}:
                return True
            if s in {"no", "off", "false", "0", "/off"}:
                return False
            return None

        acad_raw = fval("academiarecognitioncheck")
        cert_raw = fval("certifiedrecognitioncheck")
        t = checkbox_truthy(acad_raw)
        academia = "Yes" if t else ("No" if t is False else None)
        t = checkbox_truthy(cert_raw)
        certified = "Yes" if t else ("No" if t is False else None)

        contact_name = fval("contact name")
        email = fval("contact email")
        phone = fval("contact phone")
        postal = fval("postal address")

        syllabi_desc = fval("descriptino of how syllabi are integrated", "description of how syllabi are integrated")
        courses_modules = fval("list of courses and modules")
        proof = fval("proof of certifications", "proof of istqb")

        urls = fval("university website links", "website links", "university website")
        additional = fval("additional relevant information or documents")

        signature_date = fval("signature date")

        # Nová pole – z AcroForm pokud existují
        printed_name_title = fval("printed name", "name and title", "printed name, title")
        receiving_member_board = fval("receiving member board")
        date_received = fval("date received")
        validity_start_date = fval("validity start")
        validity_end_date = fval("validity end")
    else:
        additional = None

    # ---------- Textové fallbacky (pouze pokud prázdné) ----------
    import re as _re

    def _take_section(text_src: str, start_label: str, next_label: str) -> str:
        m = _re.search(rf"\b{_re.escape(start_label)}\b(.*?)(?:\b{_re.escape(next_label)}\b|\Z)", text_src, flags=_re.IGNORECASE | _re.DOTALL)
        return m.group(1) if m else ""

    # 5. Eligibility Evidence – obsah bývá v této sekci
    sec5 = _take_section(norm, "5. Eligibility Evidence", "6. Declaration")

    # University website links – URL z textu, deduplikace, stabilní pořadí
    if not urls:
        RE_URL = _re.compile(r"https?://[^\s<>()]+", _re.IGNORECASE)
        cand = sec5 if sec5 else norm
        found = RE_URL.findall(cand)
        if found:
            seen = set()
            dedup = []
            for u in found:
                if u not in seen:
                    seen.add(u)
                    dedup.append(u)
            urls = "\n".join(dedup) if dedup else None

    # Additional relevant information – text za labelem v sekci 5
    if not additional:
        block = sec5 if sec5 else norm
        m = _re.search(r"Any\s+additional\s+relevant\s+information\s+or\s+documents(?:\s*\(if any\))?\s*:\s*(.+?)(?:\n\s*\n|\b6\.|\Z)",
                       block, flags=_re.IGNORECASE | _re.DOTALL)
        if m:
            val = m.group(1).strip()
            val = _re.sub(r"\s+", " ", val).strip()
            # Ignoruj šablonový boilerplate (prázdný formulář), ne reálná data.
            if val.lower().startswith("the istqb") and "subject to the successful eligibility" in val.lower():
                val = ""
            additional = val or None

    # Printed Name, Title – opatrně, jen pokud obsahuje jméno kandidáta
    if not printed_name_title:
        cand_name = (candidate or "").strip()
        scope_m = _re.search(r"\b6\.?\s*Declaration.*", norm, flags=_re.IGNORECASE | _re.DOTALL)
        scope = scope_m.group(0) if scope_m else norm
        if cand_name:
            # rozvolněný vzor na jméno (mezery/diakritika)
            def _fuzzy_name(name: str) -> str:
                parts = name.strip().split()
                chunks = []
                for p in parts:
                    chunks.append("\s*".join([_re.escape(ch) for ch in p]))
                return "\s+".join(chunks)

            name_re = _fuzzy_name(cand_name)
            date_pat = (r"(?:\d{4}[/.-]\d{1,2}[/.-]\d{1,2}|\d{1,2}[/.-]\d{1,2}[/.-]\d{4}|[A-Za-z]+\s+\d{1,2},?\s*\d{4})")
            patt = _re.compile(rf"({name_re})\s*,\s*([A-Za-z][A-Za-z .'-]{{1,80}}?)\s*{date_pat}", flags=_re.IGNORECASE | _re.DOTALL)
            m = patt.search(scope)
            if m:
                title = _re.sub(r"\s+", " ", m.group(2)).strip()
                printed_name_title = f"{cand_name}, {title}"

    # ---------- Fallback pro flattened PDF bez AcroForm polí ----------
    # Spustí se jen když formulář nemá žádná pole; doplní pouze prázdné hodnoty.
    if not form_fields:
        nf = _extract_no_acroform_text(norm)
        if institution is None:            institution = nf["institution_name"]
        if candidate is None:              candidate = nf["candidate_name"]
        if contact_name is None:           contact_name = nf["contact_full_name"]
        if email is None:                  email = nf["contact_email"]
        if phone is None:                  phone = nf["contact_phone"]
        if postal is None:                 postal = nf["contact_postal_address"]
        if signature_date is None:         signature_date = nf["signature_date"]
        if printed_name_title is None:     printed_name_title = nf["printed_name_title"]

    # ---------- Výstup ----------
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
        "printed_name_title": printed_name_title,
        "receiving_member_board": receiving_member_board,
        "date_received": date_received,
        "validity_start_date": validity_start_date,
        "validity_end_date": validity_end_date,
    }
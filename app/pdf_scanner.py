from __future__ import annotations
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional

from .pdf_parser import read_pdf_text, read_pdf_form_fields, parse_istqb_academia_application, guess_signature_date
from .istqb_boards import KNOWN_BOARDS

@dataclass
class PdfRecord:
    board: str
    path: Path
    size_bytes: int
    application_type: Optional[str]
    institution_name: Optional[str]
    candidate_name: Optional[str]
    recognition_academia: Optional[str]
    recognition_certified: Optional[str]
    contact_full_name: Optional[str]
    contact_email: Optional[str]
    contact_phone: Optional[str]
    contact_postal_address: Optional[str]
    proof_of_istqb_certifications: Optional[str]
    syllabi_integration_description: Optional[str]
    courses_modules_list: Optional[str]
    university_links: Optional[str]
    additional_information_documents: Optional[str]
    # --- NOVÁ POLE (6 + 7) ---
    printed_name_title: Optional[str]
    signature_date: Optional[str]  # původní
    receiving_member_board: Optional[str]
    date_received: Optional[str]
    validity_start_date: Optional[str]
    validity_end_date: Optional[str]
    # ---
    board_known: bool

    def as_row(self) -> List[str]:
        # POŘADÍ MUSÍ SEDĚT S Overview HLAVIČKAMI
        return [
            self.board or "",
            self.application_type or "",
            self.institution_name or "",
            self.candidate_name or "",
            self.recognition_academia or "",
            self.recognition_certified or "",
            self.contact_full_name or "",
            self.contact_email or "",
            self.contact_phone or "",
            self.contact_postal_address or "",
            self.syllabi_integration_description or "",
            self.courses_modules_list or "",
            self.proof_of_istqb_certifications or "",
            self.university_links or "",
            self.additional_information_documents or "",
            # Consent (nově)
            self.printed_name_title or "",
            self.signature_date or "",
            # ISTQB internal (nově)
            self.receiving_member_board or "",
            self.date_received or "",
            self.validity_start_date or "",
            self.validity_end_date or "",
            # File
            self.path.name,
        ]

    def to_dict(self) -> Dict:
        d = asdict(self)
        d["path"] = str(self.path)
        return d


class PdfScanner:
    def __init__(self, root: Path) -> None:
        self.root = root

    def _parse_one(self, path: Path) -> PdfRecord:
        fields = read_pdf_form_fields(path)
        text = read_pdf_text(path)

        def fval(*keys: str) -> str | None:
            if not fields:
                return None
            for k in fields.keys():
                kk = str(k).strip().lower()
                for probe in keys:
                    if probe in kk:
                        v = fields[k]
                        # /V hodnota, případně string
                        if isinstance(v, dict):
                            val = v.get("/V") or v.get("V")
                        else:
                            val = v
                        if val is None:
                            continue
                        s = str(val).strip()
                        if s.startswith("/"):
                            s = s[1:].lstrip()
                        if s:
                            return s
            return None

        # Section 1–5 (beze změny)
        app_type = (fval("application type") or "").strip()
        if app_type.startswith('/'):
            app_type = app_type[1:].lstrip()

        institution = fval("name of university", "high", "technical school") or ""
        candidate = fval("name of candidate") or ""

        def checkbox_truthy(raw: Optional[str]) -> bool | None:
            if raw is None:
                return None
            s = raw.strip().lower()
            if s in {"yes", "on", "true", "1", "/yes"}:
                return True
            if s in {"no", "off", "false", "0", "/no"}:
                return False
            return None

        raw_acad = fval("academiarecognitioncheck", "academia recognition", "academia_recognition")
        raw_cert = fval("certifiedrecognitioncheck", "certified recognition", "certified_recognition")
        acad_yes = checkbox_truthy(raw_acad)
        cert_yes = checkbox_truthy(raw_cert)
        recog_acad = "Yes" if acad_yes else "No"
        recog_cert = "Yes" if cert_yes else "No"

        contact_name  = fval("full name", "contact name") or ""
        contact_email = fval("email") or ""
        contact_phone = fval("phone") or ""
        contact_addr  = fval("postal address") or ""

        sig_iso = guess_signature_date(fields, text) or ""

        syllabi_desc = fval("syllabi", "integrated") or ""
        courses_list = fval("courses/modules", "courses and modules", "courses") or ""
        proof_cert   = fval("proof of istqb", "proof of certifications") or ""
        uni_links    = fval("university website", "website links") or ""
        addl_info    = fval("additional relevant information", "additional information") or ""

        # --- NOVÁ POLE z parse_istqb_academia_application (raw) ---
        extra: Dict[str, Optional[str]] = {}
        try:
            extra = parse_istqb_academia_application(text or "", fields or {})
        except Exception:
            extra = {}

        printed_name_title = extra.get("printed_name_title") or ""
        rmb                 = extra.get("receiving_member_board") or ""
        date_received       = extra.get("date_received") or ""
        validity_start      = extra.get("validity_start_date") or ""
        validity_end        = extra.get("validity_end_date") or ""

        board = self._derive_board(path) if hasattr(self, "_derive_board") else (path.parent.name or "Unknown")
        return PdfRecord(
            board=board,
            path=path,
            size_bytes=path.stat().st_size,
            application_type=app_type or None,
            institution_name=institution or None,
            candidate_name=candidate or None,
            recognition_academia=recog_acad or None,
            recognition_certified=recog_cert or None,
            contact_full_name=contact_name or None,
            contact_email=contact_email or None,
            contact_phone=contact_phone or None,
            contact_postal_address=contact_addr or None,
            proof_of_istqb_certifications=proof_cert or None,
            syllabi_integration_description=syllabi_desc or None,
            courses_modules_list=courses_list or None,
            university_links=uni_links or None,
            additional_information_documents=addl_info or None,
            printed_name_title=(printed_name_title or None),
            signature_date=sig_iso or None,
            receiving_member_board=(rmb or None),
            date_received=(date_received or None),
            validity_start_date=(validity_start or None),
            validity_end_date=(validity_end or None),
            board_known=board in KNOWN_BOARDS,
        )

    def scan(self) -> List[PdfRecord]:
        records: List[PdfRecord] = []
        if self.root is None or not self.root.exists():
            return records

        for path in self.root.rglob("*"):
            try:
                if not path.is_file():
                    continue
                if path.suffix.lower() != ".pdf":
                    continue
                rel = path.relative_to(self.root)
                if "__archive__" in rel.parts:
                    continue
            except Exception:
                continue

            try:
                rec = self._parse_one(path)
                records.append(rec)
            except Exception:
                continue
        return records
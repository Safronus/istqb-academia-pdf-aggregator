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
    signature_date: Optional[str]
    proof_of_istqb_certifications: Optional[str]
    syllabi_integration_description: Optional[str]
    courses_modules_list: Optional[str]
    university_links: Optional[str]
    additional_information_documents: Optional[str]
    board_known: bool
    # NOVÉ – sekce 6 a 7:
    printed_name_title: Optional[str] = None
    istqb_receiving_board: Optional[str] = None
    istqb_date_received: Optional[str] = None
    istqb_valid_from: Optional[str] = None
    istqb_valid_to: Optional[str] = None

    def as_row(self) -> List[str]:
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
            self.signature_date or "",
            self.path.name,
        ]

    def to_dict(self) -> Dict:
        d = asdict(self)
        d["path"] = str(self.path)
        return d


class PdfScanner:
    def __init__(self, root: Path) -> None:
        self.root = root

    def _derive_board(self, pdf_path: Path) -> str:
        try:
            rel = pdf_path.relative_to(self.root)
            parts = rel.parts
            if len(parts) >= 2:
                return parts[0]
            return "Unknown"
        except Exception:
            return "Unknown"

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
                            val = str(v)
                        if val is None:
                            continue
                        s = str(val).strip()
                        if s:
                            return s
            return None

        # Application type (radio)
        app_type = fval("application type") or ""
        app_type = app_type.strip()
        
        if app_type.startswith('/'):
            app_type = app_type[1:].lstrip()  # v0.6c: remove leading slash from PDF Name

        # Institution & candidate (sekce 2)
        institution = fval("name of university", "high", "technical school") or ""
        candidate = fval("name of candidate") or ""

        # Recognitions (sekce 3)
        def checkbox_truthy(val: object) -> bool:
            if val is None:
                return False
            s = str(val).strip().lower()
            # běžné návraty z pdf: '/yes', '/off', 'on', 'off', 'true', 'false', '1', '0', 'checked'
            return s in {"yes", "/yes", "on", "true", "1", "checked", "selected"}

        # vezmi první nalezenou hodnotu ve form fields
        raw_acad = fval("academiarecognitioncheck", "academia recognition", "academia_recognition")
        raw_cert = fval("certifiedrecognitioncheck", "certified recognition", "certified_recognition")

        acad_yes = checkbox_truthy(raw_acad)
        cert_yes = checkbox_truthy(raw_cert)

        recog_acad = "Yes" if acad_yes else "No"
        recog_cert = "Yes" if cert_yes else "No"

        # Contacts (sekce 4)
        contact_name  = fval("full name", "contact name") or ""
        contact_email = fval("email") or ""
        contact_phone = fval("phone") or ""
        contact_addr  = fval("postal address") or ""

        # Signature date – form fields / text → vždy YYYY-MM-DD
        sig_iso = guess_signature_date(fields, text) or ""

        # Eligibility (sekce 5) – načteno, v Overview skryto
        syllabi_desc = fval("syllabi", "integrated") or ""
        courses_list = fval("courses/modules", "courses and modules", "courses") or ""
        proof_cert   = fval("proof of istqb", "proof of certifications") or ""
        uni_links    = fval("university website", "website links") or ""
        addl_info    = fval("additional relevant information", "additional information") or ""

        board = self._derive_board(path)
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
            signature_date=sig_iso or None,
            proof_of_istqb_certifications=proof_cert or None,
            syllabi_integration_description=syllabi_desc or None,
            courses_modules_list=courses_list or None,
            university_links=uni_links or None,
            additional_information_documents=addl_info or None,
            board_known=board in KNOWN_BOARDS,
        )

    def scan(self) -> List[PdfRecord]:
        records: List[PdfRecord] = []
        if self.root is None or not self.root.exists():
            return records
    
        # Projdeme všechny PDF (case-insensitive), ale IGNORUJEME podsložky "__archive__"
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
                # tiché přeskočení, abychom nezastavili sken
                continue
    
        return records
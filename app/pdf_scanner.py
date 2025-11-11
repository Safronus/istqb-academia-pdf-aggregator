from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional

from .pdf_parser import read_pdf_text, read_pdf_form_fields, parse_istqb_academia_application
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

        # Institution & candidate (sekce 2)
        institution = fval("name of university", "high", "technical school") or ""
        candidate = fval("name of candidate") or ""

        # Recognitions (sekce 3)
        recog_acad = fval("academiarecognitioncheck") or fval("academia recognition") or ""
        recog_cert = fval("certifiedrecognitioncheck") or fval("certified recognition") or ""
        # Normalizace přepínačů Ano/Off → Yes/No
        def norm_check(s: str) -> str:
            s = s.strip().lower()
            return "Yes" if s in ("yes", "on", "true", "1", "checked") else ("No" if s in ("no", "off", "false", "0") else s or "")
        recog_acad = norm_check(recog_acad)
        recog_cert = norm_check(recog_cert)

        # Contacts (sekce 4)
        contact_name  = fval("full name", "contact name") or ""
        contact_email = fval("email") or ""
        contact_phone = fval("phone") or ""
        contact_addr  = fval("postal address") or ""

        # Signature date – klíčová část: form fields → text; vždy YYYY-MM-DD
        sig_iso = guess_signature_date(fields, text) or ""

        # Eligibility (sekce 5) – načítáme, ale v Overview skryjeme
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
        """Recursively find all PDF files under root and parse records.
        Case-insensitive handling of '.pdf' vs '.PDF' etc.
        """
        records: List[PdfRecord] = []
        if not self.root.exists():
            return records

        for path in self.root.rglob("*"):
            if not path.is_file():
                continue
            if path.suffix.lower() != ".pdf":
                continue
            try:
                records.append(self._parse_one(path))
            except Exception:
                board = self._derive_board(path)
                records.append(
                    PdfRecord(
                        board=board,
                        path=path,
                        size_bytes=path.stat().st_size,
                        application_type=None,
                        institution_name=None,
                        candidate_name=None,
                        recognition_academia=None,
                        recognition_certified=None,
                        contact_full_name=None,
                        contact_email=None,
                        contact_phone=None,
                        contact_postal_address=None,
                        signature_date=None,
                        proof_of_istqb_certifications=None,
                        syllabi_integration_description=None,
                        courses_modules_list=None,
                        university_links=None,
                        additional_information_documents=None,
                        board_known=board in KNOWN_BOARDS,
                    )
                )
        return records
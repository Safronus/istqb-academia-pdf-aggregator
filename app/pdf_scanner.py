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
    # New fields (Section 5)
    syllabi_integration_description: Optional[str]
    courses_modules_list: Optional[str]
    university_links: Optional[str]
    additional_information_documents: Optional[str]
    board_known: bool

    def as_row(self) -> List[str]:
        # Overview table remains unchanged intentionally (minimal-change);
        # we keep Proof short and include file path.
        return [
            self.board,
            self.application_type or "",
            self.institution_name or "",
            self.candidate_name or "",
            self.recognition_academia or "",
            self.recognition_certified or "",
            self.contact_full_name or "",
            self.contact_email or "",
            self.contact_phone or "",
            self.contact_postal_address or "",
            self.signature_date or "",
            (self.proof_of_istqb_certifications or "")[:120].strip(),
            str(self.path),
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

    def _parse_one(self, pdf_path: Path) -> PdfRecord:
        text = read_pdf_text(pdf_path)
        form_fields = read_pdf_form_fields(pdf_path)
        fields = parse_istqb_academia_application(text, form_fields=form_fields)
        board = self._derive_board(pdf_path)
        return PdfRecord(
            board=board,
            path=pdf_path,
            size_bytes=pdf_path.stat().st_size,
            application_type=fields.get("application_type"),
            institution_name=fields.get("institution_name"),
            candidate_name=fields.get("candidate_name"),
            recognition_academia=fields.get("recognition_academia"),
            recognition_certified=fields.get("recognition_certified"),
            contact_full_name=fields.get("contact_full_name"),
            contact_email=fields.get("contact_email"),
            contact_phone=fields.get("contact_phone"),
            contact_postal_address=fields.get("contact_postal_address"),
            signature_date=fields.get("signature_date"),
            proof_of_istqb_certifications=fields.get("proof_of_istqb_certifications"),
            syllabi_integration_description=fields.get("syllabi_integration_description"),
            courses_modules_list=fields.get("courses_modules_list"),
            university_links=fields.get("university_links"),
            additional_information_documents=fields.get("additional_information_documents"),
            board_known=board in KNOWN_BOARDS,
        )

    def scan(self) -> List[PdfRecord]:
        records: List[PdfRecord] = []
        if not self.root.exists():
            return records
        for path in self.root.rglob("*.pdf"):
            if not path.is_file():
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
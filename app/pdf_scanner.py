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
    syllabi_integration_description: Optional[str]
    courses_modules_list: Optional[str]
    proof_of_istqb_certifications: Optional[str]
    university_links: Optional[str]
    additional_information_documents: Optional[str]
    # Consent
    printed_name_title: Optional[str]              # NOVĚ
    signature_date: Optional[str]
    # ISTQB internal (Use Only)
    receiving_member_board: Optional[str]          # NOVĚ
    date_received: Optional[str]                   # NOVĚ
    validity_start_date: Optional[str]             # NOVĚ
    validity_end_date: Optional[str]               # NOVĚ
    # Meta
    board_known: bool

    def as_row(self) -> List[str]:
        """
        Pořadí MUSÍ odpovídat hlavičkám v Overview (+ File name, Sorted).
        """
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
            # NOVÝ sloupec před Signature Date
            self.printed_name_title or "",
            # Původní Signature Date
            self.signature_date or "",
            # NOVÉ 4 sloupce za Signature Date
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

        parsed = parse_istqb_academia_application(text, fields) if text is not None else {}

        # Fallback pro datum podpisu, pokud parser nic nenašel
        sig = parsed.get("signature_date")
        if not sig:
            sig = guess_signature_date(fields or {}, text or "")

        # Board + známý board
        board = self._derive_board(path)
        known = (board in KNOWN_BOARDS)

        # Vytvoř záznam
        return PdfRecord(
            board=board,
            path=path,
            size_bytes=path.stat().st_size if path.exists() else 0,
            application_type=parsed.get("application_type"),
            institution_name=parsed.get("institution_name"),
            candidate_name=parsed.get("candidate_name"),
            recognition_academia=parsed.get("recognition_academia"),
            recognition_certified=parsed.get("recognition_certified"),
            contact_full_name=parsed.get("contact_full_name"),
            contact_email=parsed.get("contact_email"),
            contact_phone=parsed.get("contact_phone"),
            contact_postal_address=parsed.get("contact_postal_address"),
            syllabi_integration_description=parsed.get("syllabi_integration_description"),
            courses_modules_list=parsed.get("courses_modules_list"),
            proof_of_istqb_certifications=parsed.get("proof_of_istqb_certifications"),
            university_links=parsed.get("university_links"),
            additional_information_documents=parsed.get("additional_information_documents"),
            printed_name_title=parsed.get("printed_name_title"),
            signature_date=sig,
            receiving_member_board=parsed.get("receiving_member_board"),
            date_received=parsed.get("date_received"),
            validity_start_date=parsed.get("validity_start_date"),
            validity_end_date=parsed.get("validity_end_date"),
            board_known=known,
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
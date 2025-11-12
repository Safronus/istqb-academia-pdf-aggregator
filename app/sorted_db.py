from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional


class SortedDb:
    """
    JSON DB asociovaná se složkou 'Sorted PDFs'.
    Struktura:
    {
      "version": "1.0",
      "updated": "ISO8601",
      "records": {
        "<rel_path_from_sorted_root>": {
          "board": "CaSQB",
          "file_name": "XYZ.pdf",
          "edited": false,
          "created": "ISO8601",
          "updated": "ISO8601",
          "data": { ... extrahovaná pole ... }
        },
        ...
      }
    }
    """

    def __init__(self, sorted_root: Path, db_name: str = "sorted_db.json") -> None:
        self.sorted_root = Path(sorted_root)
        self.db_path = self.sorted_root / db_name
        self.doc: Dict[str, Any] = {"version": "1.0", "updated": None, "records": {}}

    def load(self) -> None:
        try:
            if self.db_path.exists():
                self.doc = json.loads(self.db_path.read_text(encoding="utf-8"))
            else:
                self._touch()
        except Exception:
            # poškozený soubor -> založ nový (bezpečný fallback)
            self.doc = {"version": "1.0", "updated": None, "records": {}}
            self._touch()

    def save(self) -> None:
        """
        Uloží obsah JSON DB na disk. Zároveň bezpečně serializuje
        nativní Python objekty, které JSON neumí (např. Path).
        Minimal-change: pouze serializace při zápisu, struktura self.doc se jinak nemění.
        """
        from pathlib import Path as _Path
        from dataclasses import is_dataclass as _is_dc, asdict as _asdict
    
        def _default(o):
            # Nejčastější případ: pathlib.Path
            if isinstance(o, _Path):
                return str(o)
            # Případné dataclass zůstanou čitelné v JSON
            if _is_dc(o):
                return _asdict(o)
            # Konzervativní doplněk: set -> list (pro jistotu)
            if isinstance(o, set):
                return list(o)
            # Poslední záchrana: převod na str (minimal-change)
            return str(o)
    
        self.doc["updated"] = datetime.utcnow().isoformat()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.db_path.write_text(
            json.dumps(self.doc, ensure_ascii=False, indent=2, default=_default),
            encoding="utf-8",
        )

    def _touch(self) -> None:
        self.save()

    def key_for(self, abs_path: Path) -> str:
        """Relativní klíč v rámci sorted_root (POSIX)."""
        rel = Path(abs_path).resolve().relative_to(self.sorted_root.resolve())
        return rel.as_posix()

    def get(self, abs_path: Path) -> Optional[Dict[str, Any]]:
        key = self.key_for(abs_path)
        return self.doc.get("records", {}).get(key)

    def upsert_parsed(self, abs_path: Path, board: str, file_name: str, data: Dict[str, Any]) -> None:
        """
        Vloží/aktualizuje záznam z parsingu.
        Pokud záznam existuje a 'edited' == True, data NEPŘEPISUJE (zůstává editovaný stav).
        Jinak data aktualizuje a nastaví edited=False.
        """
        key = self.key_for(abs_path)
        now = datetime.utcnow().isoformat()
        recs: Dict[str, Any] = self.doc.setdefault("records", {})
        existing = recs.get(key)
        if existing and existing.get("edited"):
            # respektuj ruční editaci – pouze dotkni metadata
            existing["updated"] = now
            if board:
                existing["board"] = board
            if file_name:
                existing["file_name"] = file_name
            return

        recs[key] = {
            "board": board,
            "file_name": file_name,
            "edited": False,
            "created": existing.get("created") if existing else now,
            "updated": now,
            "data": data,
        }

    def mark_edited(self, abs_path: Path, new_data: Dict[str, Any]) -> None:
        """Uloží ruční úpravu dat a nastaví edited=True."""
        key = self.key_for(abs_path)
        now = datetime.utcnow().isoformat()
        recs: Dict[str, Any] = self.doc.setdefault("records", {})
        rec = recs.setdefault(key, {"created": now})
        rec["updated"] = now
        rec["edited"] = True
        rec["data"] = new_data

    def iter_items(self):
        for key, rec in self.doc.get("records", {}).items():
            yield key, rec
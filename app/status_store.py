from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Dict

from PySide6.QtCore import QStandardPaths

# Workflow states (order matters for UI). First entry is the default.
STATUSES = [
    "In Progress",
    "Completed",
    "Ready for Web",
    "Published on Web",
    "Problematic",
]
DEFAULT_STATUS = STATUSES[0]


class StatusStore:
    """Persists a per-PDF workflow status, keyed by a stable relative key
    (typically '<board>/<file_name>'). Stored as a local JSON in the OS config
    dir (never in the project tree, never versioned)."""

    APP_DIR_NAME = "istqb-academia-aggregator"
    FILE_NAME = "statuses.json"

    def __init__(self) -> None:
        base = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
        if not base:
            base = str(Path.home() / ".config")
        self.path = Path(base) / self.APP_DIR_NAME / self.FILE_NAME
        self.doc: Dict[str, dict] = {"version": 1, "statuses": {}}

    def load(self) -> None:
        try:
            if self.path.exists():
                loaded = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict) and isinstance(loaded.get("statuses"), dict):
                    self.doc = loaded
        except Exception:
            self.doc = {"version": 1, "statuses": {}}

    def save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(
                json.dumps(self.doc, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def get(self, key: str) -> str:
        if not key:
            return DEFAULT_STATUS
        rec = self.doc.get("statuses", {}).get(key)
        if isinstance(rec, dict):
            st = rec.get("status")
            if st in STATUSES:
                return st
        return DEFAULT_STATUS

    def set(self, key: str, status: str) -> None:
        if not key or status not in STATUSES:
            return
        self.doc.setdefault("statuses", {})[key] = {
            "status": status,
            "updated": datetime.utcnow().isoformat(),
        }

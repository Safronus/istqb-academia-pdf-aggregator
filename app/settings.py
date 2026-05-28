from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from PySide6.QtCore import QStandardPaths


_DEFAULTS: Dict[str, Any] = {
    "pdf_root": None,
    "sorted_root": None,
    "window_geometry": None,   # base64 string of QMainWindow.saveGeometry()
    "active_tab": 0,
    "filters": {
        "overview_search": "",
        "overview_board": "",
        "hide_sorted": True,
    },
}


class AppSettings:
    """Persisted application settings as a human-readable JSON file.

    Stored in the platform config dir (outside the project tree) so it never
    mixes with source code or applicant data.
    """

    APP_DIR_NAME = "istqb-academia-aggregator"
    FILE_NAME = "settings.json"

    def __init__(self) -> None:
        base = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
        if not base:
            base = str(Path.home() / ".config")
        self.path = Path(base) / self.APP_DIR_NAME / self.FILE_NAME
        # deep copy of defaults
        self.data: Dict[str, Any] = json.loads(json.dumps(_DEFAULTS))

    def load(self) -> None:
        try:
            if self.path.exists():
                loaded = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    for k, v in loaded.items():
                        if k == "filters" and isinstance(v, dict):
                            self.data["filters"].update(v)
                        else:
                            self.data[k] = v
        except Exception:
            # corrupt/unreadable file -> keep defaults
            pass

    def save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(
                json.dumps(self.data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value

    def get_filter(self, key: str, default: Any = None) -> Any:
        return self.data.get("filters", {}).get(key, default)

    def set_filter(self, key: str, value: Any) -> None:
        self.data.setdefault("filters", {})[key] = value

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
from app.main_window import MainWindow
from app.themes import apply_dark_palette


def resolve_pdf_root(arg: Optional[str]) -> Path:
    if arg:
        return Path(arg).expanduser().resolve()
    return (Path(__file__).parent / "PDF").resolve()


def main() -> None:
    # HiDPI / Retina
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    apply_dark_palette(app)

    # Optional CLI: --pdf-root /path/to/folder
    pdf_root: Optional[str] = None
    if "--pdf-root" in sys.argv:
        idx = sys.argv.index("--pdf-root")
        if idx + 1 < len(sys.argv):
            pdf_root = sys.argv[idx + 1]

    root = resolve_pdf_root(pdf_root)
    root.mkdir(parents=True, exist_ok=True)

    win = MainWindow(pdf_root=root)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
from app.main_window import MainWindow
from app.themes import apply_dark_palette  # FIX: ve 0.3h je apply_dark_palette

def default_pdf_root() -> Path:
    return (Path(__file__).parent / "PDF").resolve()

def main() -> None:
    app = QApplication(sys.argv)

    # HiDPI (původní chování)
    try:
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    except Exception:
        pass

    # Dark palette (správná funkce z themes.py)
    apply_dark_palette(app)

    # Volitelný argument --pdf-root
    pdf_root_arg: Optional[str] = None
    if "--pdf-root" in sys.argv:
        try:
            idx = sys.argv.index("--pdf-root")
            if idx + 1 < len(sys.argv):
                pdf_root_arg = sys.argv[idx + 1]
        except Exception:
            pdf_root_arg = None

    cli_pdf_root = Path(pdf_root_arg).expanduser().resolve() if pdf_root_arg else None

    win = MainWindow(default_pdf_root=default_pdf_root(), cli_pdf_root=cli_pdf_root)

    # Pokud je v projektu nainstalován "manual refresh guard", aktivuj ho
    try:
        if hasattr(win, "install_manual_refresh_guard"):
            win.install_manual_refresh_guard()
    except Exception:
        pass

    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
from __future__ import annotations
import sys
from pathlib import Path
from PySide6.QtWidgets import QApplication
from app.main_window import MainWindow
from app.themes import apply_dark_theme

def main() -> None:
    app = QApplication(sys.argv)

    # HiDPI (původní nastavení ve tvém projektu – pokud máš jinde, ponech)
    try:
        from PySide6.QtCore import Qt
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    except Exception:
        pass

    apply_dark_theme(app)

    # PDF root volitelně podle projektu; v 0.3h bývá Path.cwd()/PDF
    pdf_root = Path.cwd() / "PDF"
    win = MainWindow(pdf_root if pdf_root.exists() else None)

    # >>> DOPLNĚNO: vypni real-time obnovu a povol jen manuální refresh
    try:
        win.install_manual_refresh_guard()
    except Exception:
        pass
    # <<<

    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
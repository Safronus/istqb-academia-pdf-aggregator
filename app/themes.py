from __future__ import annotations

from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import QApplication


def apply_dark_palette(app: QApplication) -> None:
    """Apply a pleasant dark Fusion-like palette."""
    palette = QPalette()

    # Base colors
    bg = QColor(30, 30, 30)
    base = QColor(36, 36, 36)
    alt = QColor(45, 45, 45)
    text = QColor(220, 220, 220)
    disabled = QColor(127, 127, 127)
    highlight = QColor(53, 132, 228)

    palette.setColor(QPalette.Window, bg)
    palette.setColor(QPalette.WindowText, text)
    palette.setColor(QPalette.Base, base)
    palette.setColor(QPalette.AlternateBase, alt)
    palette.setColor(QPalette.ToolTipBase, base)
    palette.setColor(QPalette.ToolTipText, text)
    palette.setColor(QPalette.Text, text)
    palette.setColor(QPalette.Button, alt)
    palette.setColor(QPalette.ButtonText, text)
    palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.Highlight, highlight)
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    palette.setColor(QPalette.PlaceholderText, disabled)

    app.setPalette(palette)
    app.setStyleSheet("""
        QLineEdit, QComboBox, QTextEdit, QTreeView, QTableView {
            selection-background-color: #3584e4;
            gridline-color: #404040;
            color: #e0e0e0;
            background: #242424;
        }
        QHeaderView::section {
            background: #2d2d2d;
            color: #e0e0e0;
            padding: 4px 6px;
            border: 0px;
        }
        QPushButton {
            background: #2d2d2d;
            border: 1px solid #3c3c3c;
            padding: 6px 10px;
            border-radius: 8px;
        }
        QPushButton:hover { border-color: #5a5a5a; }
    """)
from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import QSortFilterProxyModel, Qt, QItemSelectionModel, QModelIndex
from PySide6.QtGui import QAction, QDesktopServices
from PySide6.QtWidgets import (
    QWidget, QMainWindow, QTabWidget, QVBoxLayout, QHBoxLayout,
    QTableView, QLineEdit, QPushButton, QLabel, QComboBox, QFormLayout,
    QTreeView, QFileSystemModel, QSplitter, QMessageBox
)

from .pdf_scanner import PdfScanner, PdfRecord
from .istqb_boards import KNOWN_BOARDS


class RecordsModel(QSortFilterProxyModel):
    def __init__(self, headers: List[str], parent=None):
        super().__init__(parent)
        self.headers = headers
        self.search = ""
        self.board_filter = "All"

    def set_search(self, text: str) -> None:
        self.search = text.lower().strip()
        self.invalidateFilter()

    def set_board(self, board: str) -> None:
        self.board_filter = board
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        model = self.sourceModel()
        if model is None:
            return True
        idx_board = model.index(source_row, 0, source_parent)
        idx_all_text = [model.index(source_row, c, source_parent) for c in range(model.columnCount())]

        board_val = model.data(idx_board, Qt.DisplayRole) or ""
        if self.board_filter != "All" and board_val != self.board_filter:
            return False

        if not self.search:
            return True

        # search across all visible columns
        for idx in idx_all_text:
            val = (model.data(idx, Qt.DisplayRole) or "").lower()
            if self.search in val:
                return True
        return False


class MainWindow(QMainWindow):
    def __init__(self, pdf_root: Path):
        super().__init__()
        self.setWindowTitle("ISTQB Academia PDF Aggregator")
        self.resize(1200, 800)
        self.pdf_root = pdf_root

        self.records: List[PdfRecord] = []

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.overview_tab = QWidget()
        self.browser_tab = QWidget()
        self.tabs.addTab(self.overview_tab, "Overview")
        self.tabs.addTab(self.browser_tab, "PDF Browser")

        self._build_menu()
        self._build_overview_tab()
        self._build_browser_tab()

        self.rescan()

    # ----- Menu / actions -----
    def _build_menu(self) -> None:
        rescan_action = QAction("Rescan", self)
        rescan_action.triggered.connect(self.rescan)

        open_action = QAction("Open Selected PDF", self)
        open_action.triggered.connect(self.open_selected_pdf)

        export_csv_action = QAction("Export CSV (visible rows)…", self)
        export_csv_action.triggered.connect(self.export_csv)

        export_xlsx_action = QAction("Export XLSX (visible rows)…", self)
        export_xlsx_action.triggered.connect(self.export_xlsx)

        about_action = QAction("About", self)
        about_action.triggered.connect(self._about)

        # Minimal-change: actions directly on menubar (no new menus introduced)
        self.menuBar().addAction(rescan_action)
        self.menuBar().addAction(open_action)
        self.menuBar().addAction(export_csv_action)
        self.menuBar().addAction(export_xlsx_action)
        self.menuBar().addAction(about_action)
        
    def _gather_visible_records(self) -> list[PdfRecord]:
        """Return records corresponding to the currently visible rows in the table."""
        model = self.table.model()
        if model is None:
            return []
        paths: list[str] = []
        if isinstance(model, QSortFilterProxyModel):
            source = model.sourceModel()
            if source is None:
                return []
            for r in range(model.rowCount()):
                # Column 12 = File (path)
                src_idx = model.mapToSource(model.index(r, 12))
                path_str = source.index(src_idx.row(), 12).data()
                if path_str:
                    paths.append(str(path_str))
        else:
            for r in range(model.rowCount()):
                path_str = model.index(r, 12).data()
                if path_str:
                    paths.append(str(path_str))

        out: list[PdfRecord] = []
        for p in paths:
            rec = next((x for x in self.records if str(x.path) == p), None)
            if rec:
                out.append(rec)
        return out
    
    def export_csv(self) -> None:
        from PySide6.QtWidgets import QFileDialog
        import csv

        records = self._gather_visible_records()
        if not records:
            QMessageBox.information(self, "Export CSV", "No rows to export (check filters/search).")
            return

        default_name = str((self.pdf_root / "export.csv"))
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", default_name, "CSV Files (*.csv)")
        if not path:
            return

        headers = [
            "Board", "Application Type", "Institution Name", "Candidate Name",
            "Recognition Academia", "Recognition Certified",
            "Contact Name", "Email", "Phone", "Postal Address",
            "Signature Date", "Proof of ISTQB Certifications", "File"
        ]
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for r in records:
                    writer.writerow(r.as_row())
            QMessageBox.information(self, "Export CSV", f"Exported {len(records)} rows to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export CSV", f"Failed to write CSV:\n{e}")
            
    def export_xlsx(self) -> None:
        from PySide6.QtWidgets import QFileDialog
        try:
            from openpyxl import Workbook
        except Exception:
            QMessageBox.critical(
                self, "Export XLSX",
                "The 'openpyxl' package is required for XLSX export.\n"
                "Install it in your environment:\n\npip install openpyxl"
            )
            return

        records = self._gather_visible_records()
        if not records:
            QMessageBox.information(self, "Export XLSX", "No rows to export (check filters/search).")
            return

        default_name = str((self.pdf_root / "export.xlsx"))
        path, _ = QFileDialog.getSaveFileName(self, "Save XLSX", default_name, "Excel Workbook (*.xlsx)")
        if not path:
            return

        headers = [
            "Board", "Application Type", "Institution Name", "Candidate Name",
            "Recognition Academia", "Recognition Certified",
            "Contact Name", "Email", "Phone", "Postal Address",
            "Signature Date", "Proof of ISTQB Certifications", "File"
        ]

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "ISTQB Applications"
            ws.append(headers)
            for r in records:
                ws.append(r.as_row())
            # Auto width (simple heuristic)
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(12, max_len + 2), 60)
            wb.save(path)
            QMessageBox.information(self, "Export XLSX", f"Exported {len(records)} rows to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export XLSX", f"Failed to write XLSX:\n{e}")

    def _about(self) -> None:
        QMessageBox.information(
            self, "About",
            "ISTQB Academia PDF Aggregator\n\n"
            "Scans the 'PDF' folder next to the script and aggregates key fields.\n"
            "UI: English • Theme: Dark • Target: macOS"
        )

    # ----- Overview tab -----
    def _build_overview_tab(self) -> None:
        layout = QVBoxLayout()
        controls = QHBoxLayout()

        self.board_combo = QComboBox()
        self.board_combo.addItem("All")
        for b in sorted(KNOWN_BOARDS):
            self.board_combo.addItem(b)
        self.board_combo.currentTextChanged.connect(self._filter_board)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search…")
        self.search_edit.textChanged.connect(self._filter_text)

        self.open_btn = QPushButton("Open PDF")
        self.open_btn.clicked.connect(self.open_selected_pdf)

        controls.addWidget(QLabel("Board:"))
        controls.addWidget(self.board_combo, 1)
        controls.addSpacing(12)
        controls.addWidget(QLabel("Search:"))
        controls.addWidget(self.search_edit, 4)
        controls.addSpacing(12)
        controls.addWidget(self.open_btn)

        layout.addLayout(controls)

        self.table = QTableView()
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.doubleClicked.connect(self.open_selected_pdf)

        layout.addWidget(self.table, 1)
        self.overview_tab.setLayout(layout)

    def _filter_board(self, txt: str) -> None:
        proxy = self.table.model()
        if isinstance(proxy, RecordsModel):
            proxy.set_board(txt)

    def _filter_text(self, txt: str) -> None:
        proxy = self.table.model()
        if isinstance(proxy, RecordsModel):
            proxy.set_search(txt)

    # ----- Browser tab -----
    def _build_browser_tab(self) -> None:
        splitter = QSplitter()
        left = QWidget()
        right = QWidget()
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        # Left: filtered tree for PDFs only
        vleft = QVBoxLayout(left)
        self.fs_model = QFileSystemModel(self)
        self.fs_model.setRootPath(str(self.pdf_root))
        self.fs_model.setNameFilterDisables(False)  # hide non-matching files
        self.fs_model.setNameFilters(["*.pdf", "*.PDF"])

        self.tree = QTreeView()
        self.tree.setModel(self.fs_model)
        self.tree.setRootIndex(self.fs_model.index(str(self.pdf_root)))
        self.tree.setSortingEnabled(True)
        self.tree.doubleClicked.connect(self._open_from_tree)
        self.tree.selectionModel().selectionChanged.connect(self._tree_selection_changed)
        vleft.addWidget(self.tree)

        # Right: details
        vright = QVBoxLayout(right)
        self.detail_form = QFormLayout()
        self.lbl_board = QLabel("-")
        self.lbl_known = QLabel("-")
        self.lbl_app_type = QLabel("-")
        self.lbl_inst = QLabel("-")
        self.lbl_cand = QLabel("-")
        self.lbl_acad = QLabel("-")
        self.lbl_cert = QLabel("-")
        self.lbl_contact = QLabel("-")
        self.lbl_email = QLabel("-")
        self.lbl_phone = QLabel("-")
        self.lbl_postal = QLabel("-")
        self.lbl_date = QLabel("-")
        self.lbl_proof = QLabel("-")
        self.btn_open_right = QPushButton("Open PDF")
        self.btn_open_right.clicked.connect(self._open_selected_detail)

        self.detail_form.addRow("Board:", self.lbl_board)
        self.detail_form.addRow("Board known:", self.lbl_known)
        self.detail_form.addRow("Application Type:", self.lbl_app_type)
        self.detail_form.addRow("Institution:", self.lbl_inst)
        self.detail_form.addRow("Candidate:", self.lbl_cand)
        self.detail_form.addRow("Recognition Academia:", self.lbl_acad)
        self.detail_form.addRow("Recognition Certified:", self.lbl_cert)
        self.detail_form.addRow("Contact Name:", self.lbl_contact)
        self.detail_form.addRow("Email:", self.lbl_email)
        self.detail_form.addRow("Phone:", self.lbl_phone)
        self.detail_form.addRow("Postal Address:", self.lbl_postal)
        self.detail_form.addRow("Signature Date:", self.lbl_date)
        self.detail_form.addRow("Proof of ISTQB Certifications:", self.lbl_proof)
        vright.addLayout(self.detail_form)
        vright.addStretch(1)
        vright.addWidget(self.btn_open_right)

        container = QWidget()
        lay = QVBoxLayout(container)
        lay.addWidget(splitter)
        self.browser_tab.setLayout(lay)

    # ----- Data -----
    def rescan(self) -> None:
        scanner = PdfScanner(self.pdf_root)
        self.records = scanner.scan()

        # Build table model
        from PySide6.QtGui import QStandardItemModel, QStandardItem

        headers = [
            "Board", "Application Type", "Institution Name", "Candidate Name",
            "Rec. Academia", "Rec. Certified",
            "Contact Name", "Email", "Phone", "Postal Address",
            "Signature Date", "Certifications (short)", "File"
        ]
        model = QStandardItemModel(0, len(headers), self)
        model.setHorizontalHeaderLabels(headers)
        for rec in self.records:
            items = [QStandardItem(cell) for cell in rec.as_row()]
            for it in items:
                it.setEditable(False)
            model.appendRow(items)

        proxy = RecordsModel(headers, self)
        proxy.setSourceModel(model)
        self.table.setModel(proxy)
        self.table.resizeColumnsToContents()
        self.table.setColumnHidden(12, False)  # show file path for clarity

    # ----- Actions -----
    def _selected_record(self) -> Optional[PdfRecord]:
        sel = self.table.selectionModel()
        if not sel or not sel.hasSelection():
            return None
        index = sel.selectedRows()[0]
        # Map proxy -> source
        proxy = self.table.model()
        if isinstance(proxy, QSortFilterProxyModel):
            index = proxy.mapToSource(index)
            model = proxy.sourceModel()
        else:
            model = self.table.model()
        path_str = model.index(index.row(), 12).data()  # File column
        for r in self.records:
            if str(r.path) == path_str:
                return r
        return None

    def open_selected_pdf(self) -> None:
        rec = self._selected_record()
        if not rec:
            QMessageBox.information(self, "Open PDF", "Please select a row first.")
            return
        QDesktopServices.openUrl(rec.path.as_uri())

    def _open_from_tree(self, index) -> None:
        path = Path(self.fs_model.filePath(index))
        if path.is_file() and path.suffix.lower() == ".pdf":
            QDesktopServices.openUrl(path.as_uri())

    def _open_selected_detail(self) -> None:
        # open currently selected file in the tree
        sel = self.tree.selectionModel()
        if not sel or not sel.hasSelection():
            QMessageBox.information(self, "Open PDF", "Please select a file in the tree.")
            return
        idx = sel.selectedRows()[0]
        self._open_from_tree(idx)

    def _tree_selection_changed(self, *_):
        sel = self.tree.selectionModel()
        if not sel or not sel.hasSelection():
            return
        idx = sel.selectedRows()[0]
        path = Path(self.fs_model.filePath(idx))
        if not path.is_file():
            return

        from .pdf_scanner import PdfScanner
        scanner = PdfScanner(self.pdf_root)
        rec: Optional[PdfRecord] = None
        try:
            rec = scanner._parse_one(path)
        except Exception:
            rec = None

        self._update_detail_panel(rec)

    def _update_detail_panel(self, rec: Optional[PdfRecord]) -> None:
        if rec is None:
            self.lbl_board.setText("-")
            self.lbl_known.setText("-")
            self.lbl_app_type.setText("-")
            self.lbl_inst.setText("-")
            self.lbl_cand.setText("-")
            self.lbl_acad.setText("-")
            self.lbl_cert.setText("-")
            self.lbl_contact.setText("-")
            self.lbl_email.setText("-")
            self.lbl_phone.setText("-")
            self.lbl_postal.setText("-")
            self.lbl_date.setText("-")
            self.lbl_proof.setText("-")
            return

        self.lbl_board.setText(rec.board)
        self.lbl_known.setText("Yes" if rec.board_known else "Unverified")
        self.lbl_app_type.setText(rec.application_type or "")
        self.lbl_inst.setText(rec.institution_name or "")
        self.lbl_cand.setText(rec.candidate_name or "")
        self.lbl_acad.setText(rec.recognition_academia or "")
        self.lbl_cert.setText(rec.recognition_certified or "")
        self.lbl_contact.setText(rec.contact_full_name or "")
        self.lbl_email.setText(rec.contact_email or "")
        self.lbl_phone.setText(rec.contact_phone or "")
        self.lbl_postal.setText(rec.contact_postal_address or "")
        self.lbl_date.setText(rec.signature_date or "")
        self.lbl_proof.setText(rec.proof_of_istqb_certifications or "")
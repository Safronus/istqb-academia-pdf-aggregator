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
        # Board filter is based on column 0 (Board)
        idx_board = model.index(source_row, 0, source_parent)
        board_val = (model.data(idx_board, Qt.DisplayRole) or "").strip()
        if self.board_filter != "All" and board_val != self.board_filter:
            return False

        if not self.search:
            return True

        # Search across all columns
        for c in range(model.columnCount()):
            idx = model.index(source_row, c, source_parent)
            val = (model.data(idx, Qt.DisplayRole) or "").lower()
            if self.search in val:
                return True
        return False

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        """Stable multi-key ordering: Board → Application Type → Candidate Name."""
        model = self.sourceModel()
        if model is None:
            return super().lessThan(left, right)

        def data(row: int, col: int) -> str:
            idx = model.index(row, col)
            return (model.data(idx, Qt.DisplayRole) or "").strip()

        # With first column "No." added, indices shift by +1:
        BOARD = 1
        APP = 2
        CAND = 4

        a_board = data(left.row(), BOARD).lower()
        b_board = data(right.row(), BOARD).lower()

        order = {"new application": 0, "additional recognition": 1}
        a_app = order.get(data(left.row(), APP).lower(), 2)
        b_app = order.get(data(right.row(), APP).lower(), 2)

        a_cand = data(left.row(), CAND).lower()
        b_cand = data(right.row(), CAND).lower()

        return (a_board, a_app, a_cand) < (b_board, b_app, b_cand)

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
        self._init_fs_watcher()

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
        model = self.table.model()
        if model is None:
            return []
        paths: list[str] = []
        FILE_COL = 16
        if isinstance(model, QSortFilterProxyModel):
            source = model.sourceModel()
            if source is None:
                return []
            for r in range(model.rowCount()):
                src_idx = model.mapToSource(model.index(r, FILE_COL))
                path_str = source.index(src_idx.row(), FILE_COL).data()
                if path_str:
                    paths.append(str(path_str))
        else:
            for r in range(model.rowCount()):
                path_str = model.index(r, FILE_COL).data()
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

        # Build headers from the current model (replace newlines for CSV)
        headers_list: list[str] = []
        src = self.table.model()
        if isinstance(src, QSortFilterProxyModel):
            src = src.sourceModel()
        if src:
            for c in range(src.columnCount()):
                h = src.headerData(c, Qt.Horizontal, Qt.DisplayRole) or ""
                headers_list.append(str(h).replace("\n", " • "))

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers_list)
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

        # Headers from model with newline replaced
        headers_list: list[str] = []
        src = self.table.model()
        if isinstance(src, QSortFilterProxyModel):
            src = src.sourceModel()
        if src:
            for c in range(src.columnCount()):
                h = src.headerData(c, Qt.Horizontal, Qt.DisplayRole) or ""
                headers_list.append(str(h).replace("\n", " • "))

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "ISTQB Applications"
            ws.append(headers_list)
            for r in records:
                ws.append(r.as_row())
            # Auto width
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
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setMinimumHeight(44)

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
        self.fs_model.setNameFilterDisables(False)
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
        self.lbl_syllabi = QLabel("-")
        self.lbl_courses = QLabel("-")
        self.lbl_proof = QLabel("-")
        self.lbl_links = QLabel("-")
        self.lbl_additional = QLabel("-")
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
        self.detail_form.addRow("Syllabi Integration:", self.lbl_syllabi)
        self.detail_form.addRow("Courses/Modules:", self.lbl_courses)
        self.detail_form.addRow("Proof of ISTQB Certifications:", self.lbl_proof)
        self.detail_form.addRow("University Links:", self.lbl_links)
        self.detail_form.addRow("Additional Info/Documents:", self.lbl_additional)
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

        from PySide6.QtGui import QStandardItemModel, QStandardItem, QBrush, QColor

        headers = [
            "No.",
            "Board",
            "Application\nApplication Type",
            "Name of Your Academic Institution\nInstitution Name",
            "Name of Your Academic Institution\nCandidate Name",
            "Wished Recognitions\nAcademia Recognition",
            "Wished Recognitions\nCertified Recognition",
            "Contact details for information exchange\nFull Name",
            "Contact details for information exchange\nEmail Address",
            "Contact details for information exchange\nPhone Number",
            "Contact details for information exchange\nPostal Address",
            "Eligibility Evidence\nSyllabi Integration",
            "Eligibility Evidence\nCourses/Modules",
            "Eligibility Evidence\nProof of ISTQB Certifications",
            "Eligibility Evidence\nUniversity Links",
            "Eligibility Evidence\nAdditional Info/Documents",
            "Signature\nSignature Date",
            "File\nFile name",
        ]

        model = QStandardItemModel(0, len(headers), self)
        model.setHorizontalHeaderLabels(headers)

        # Column groups (cell background colors; indices reflect "No." at col 0)
        COLS_APPLICATION = [2]
        COLS_INSTITUTION = [3, 4]
        COLS_RECOG      = [5, 6]
        COLS_CONTACT    = [7, 8, 9, 10]
        COLS_ELIG       = [11, 12, 13, 14, 15]

        BRUSH_APP   = QBrush(QColor(58, 74, 110))
        BRUSH_INST  = QBrush(QColor(74, 58, 110))
        BRUSH_RECOG = QBrush(QColor(58, 110, 82))
        BRUSH_CONT  = QBrush(QColor(110, 82, 58))
        BRUSH_ELIG  = QBrush(QColor(72, 110, 110))

        def paint_group(items: list[QStandardItem], cols: list[int], brush: QBrush) -> None:
            for c in cols:
                if 0 <= c < len(items):
                    items[c].setBackground(brush)

        for rec in self.records:
            row_vals = rec.as_row()  # 17 columns (without No.)
            # prepend "No." placeholder (will be renumbered after sort/filter)
            items = [QStandardItem("")] + [QStandardItem(v) for v in row_vals]
            for it in items:
                it.setEditable(False)

            # Color groups
            paint_group(items, COLS_APPLICATION, BRUSH_APP)
            paint_group(items, COLS_INSTITUTION, BRUSH_INST)
            paint_group(items, COLS_RECOG,      BRUSH_RECOG)
            paint_group(items, COLS_CONTACT,    BRUSH_CONT)
            paint_group(items, COLS_ELIG,       BRUSH_ELIG)

            # Store full path (hidden) to the last column item for reliable open/export
            FILE_COL = len(headers) - 1
            items[FILE_COL].setData(str(rec.path), Qt.UserRole + 1)

            model.appendRow(items)

        proxy = RecordsModel(headers, self)
        proxy.setSourceModel(model)
        self.table.setModel(proxy)

        # Centered headers, sizing, sorting
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        for c in range(len(headers)):
            self.table.resizeColumnToContents(c)
        self.table.sortByColumn(1, Qt.AscendingOrder)  # sort to trigger lessThan

        # Renumber "No." column to reflect current proxy order
        self._renumber_rows()

        # Connect to renumber on sort/filter changes
        try:
            self.table.model().layoutChanged.connect(self._renumber_rows)
            self.table.model().modelReset.connect(self._renumber_rows)
        except Exception:
            pass

        # Status bar
        try:
            self.statusBar().showMessage(f"{len(self.records)} PDF parsed • Root: {self.pdf_root}")
        except Exception:
            pass

        # Rebuild FS watchers
        self._rebuild_watch_list()

    # ----- Actions -----
    def _selected_record(self) -> Optional[PdfRecord]:
        sel = self.table.selectionModel()
        if not sel or not sel.hasSelection():
            return None
        index = sel.selectedRows()[0]
        proxy = self.table.model()
        FILE_COL = 16
        if isinstance(proxy, QSortFilterProxyModel):
            index = proxy.mapToSource(index)
            model = proxy.sourceModel()
        else:
            model = self.table.model()
        path_str = model.index(index.row(), FILE_COL).data()
        for r in self.records:
            if str(r.path) == path_str:
                return r
        return None
    
    def _init_fs_watcher(self) -> None:
        from PySide6.QtCore import QTimer
        from PySide6.QtCore import QFileSystemWatcher

        self._watcher = QFileSystemWatcher(self)
        self._watcher.directoryChanged.connect(self._on_fs_changed)
        self._watcher.fileChanged.connect(self._on_fs_changed)

        self._fs_debounce = QTimer(self)
        self._fs_debounce.setSingleShot(True)
        self._fs_debounce.setInterval(300)
        self._fs_debounce.timeout.connect(self._fs_debounced)

        self._rebuild_watch_list()

    def _rebuild_watch_list(self) -> None:
        """Watch PDF root, all subdirs, and all .pdf files."""
        if not hasattr(self, "_watcher"):
            return
        watcher = self._watcher
        # Clear
        try:
            if watcher.files():
                watcher.removePaths(watcher.files())
            if watcher.directories():
                watcher.removePaths(watcher.directories())
        except Exception:
            pass

        dirs = set()
        files = set()
        if self.pdf_root.exists():
            dirs.add(str(self.pdf_root))
            for p in self.pdf_root.rglob("*"):
                if p.is_dir():
                    dirs.add(str(p))
                elif p.suffix.lower() == ".pdf":
                    files.add(str(p))
        if dirs:
            watcher.addPaths(sorted(dirs))
        if files:
            watcher.addPaths(sorted(files))

    def _on_fs_changed(self, _path: str) -> None:
        # Debounce frequent events
        if hasattr(self, "_fs_debounce"):
            self._fs_debounce.start()

    def _fs_debounced(self) -> None:
        # Rescan and refresh watchers (new subdirs/files)
        self.rescan()
        self._rebuild_watch_list()

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
            for lbl in (
                self.lbl_board, self.lbl_known, self.lbl_app_type, self.lbl_inst, self.lbl_cand,
                self.lbl_acad, self.lbl_cert, self.lbl_contact, self.lbl_email, self.lbl_phone,
                self.lbl_postal, self.lbl_date, self.lbl_syllabi, self.lbl_courses,
                self.lbl_proof, self.lbl_links, self.lbl_additional
            ):
                lbl.setText("-")
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
        self.lbl_syllabi.setText(rec.syllabi_integration_description or "")
        self.lbl_courses.setText(rec.courses_modules_list or "")
        self.lbl_proof.setText(rec.proof_of_istqb_certifications or "")
        self.lbl_links.setText(rec.university_links or "")
        self.lbl_additional.setText(rec.additional_information_documents or "")
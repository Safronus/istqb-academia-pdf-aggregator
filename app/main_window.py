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
        """Řazení: Board → Application Type → Candidate Name."""
        model = self.sourceModel()
        if model is None:
            return super().lessThan(left, right)
    
        def data(row: int, col: int) -> str:
            idx = model.index(row, col)
            return (model.data(idx, Qt.DisplayRole) or "").strip()
    
        BOARD = 0
        APP   = 1
        CAND  = 3
    
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
        FILE_COL = 16  # poslední sloupec po odebrání číslování
        if isinstance(model, QSortFilterProxyModel):
            src = model.sourceModel()
            if src is None:
                return []
            for r in range(model.rowCount()):
                pidx = model.index(r, FILE_COL)
                sidx = model.mapToSource(pidx)
                val = src.index(sidx.row(), FILE_COL).data(Qt.UserRole + 1)
                if val:
                    paths.append(str(val))
        else:
            for r in range(model.rowCount()):
                val = model.index(r, FILE_COL).data(Qt.UserRole + 1)
                if val:
                    paths.append(str(val))

        out: list[PdfRecord] = []
        for p in paths:
            rec = next((x for x in self.records if str(x.path) == p), None)
            if rec:
                out.append(rec)
        return out

    def _selected_record(self) -> Optional[PdfRecord]:
        sel = self.table.selectionModel()
        if not sel or not sel.hasSelection():
            return None
        index = sel.selectedRows()[0]
        FILE_COL = 16
        proxy = self.table.model()
        if isinstance(proxy, QSortFilterProxyModel):
            sidx = proxy.mapToSource(proxy.index(index.row(), FILE_COL))
            src = proxy.sourceModel()
        else:
            sidx = index
            src = self.table.model()
        path_str = src.index(sidx.row(), FILE_COL).data(Qt.UserRole + 1)
        for r in self.records:
            if str(r.path) == path_str:
                return r
        return None

    def _visible_columns(self) -> list[int]:
        """Return list of column indices currently visible in the table (proxy)."""
        model = self.table.model()
        if model is None:
            return []
        cols = model.columnCount()
        visible = []
        for c in range(cols):
            if not self.table.isColumnHidden(c):
                visible.append(c)
        return visible

    def export_csv(self) -> None:
        from PySide6.QtWidgets import QFileDialog
        import csv

        model = self.table.model()
        if model is None:
            QMessageBox.information(self, "Export CSV", "No data to export.")
            return

        visible_cols = self._visible_columns()
        if not visible_cols:
            QMessageBox.information(self, "Export CSV", "No visible columns to export.")
            return

        default_name = str((self.pdf_root / "export.csv"))
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", default_name, "CSV Files (*.csv)")
        if not path:
            return

        # Build headers from visible columns (display text, newlines replaced)
        headers_list: list[str] = []
        for c in visible_cols:
            h = model.headerData(c, Qt.Horizontal, Qt.DisplayRole) or ""
            headers_list.append(str(h).replace("\n", " • "))

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(headers_list)
                for r in range(model.rowCount()):
                    row_vals = []
                    for c in visible_cols:
                        val = model.index(r, c).data(Qt.DisplayRole)
                        row_vals.append("" if val is None else str(val))
                    writer.writerow(row_vals)
            QMessageBox.information(self, "Export CSV", f"Exported {model.rowCount()} rows to:\n{path}")
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

        model = self.table.model()
        if model is None:
            QMessageBox.information(self, "Export XLSX", "No data to export.")
            return

        visible_cols = self._visible_columns()
        if not visible_cols:
            QMessageBox.information(self, "Export XLSX", "No visible columns to export.")
            return

        default_name = str((self.pdf_root / "export.xlsx"))
        path, _ = QFileDialog.getSaveFileName(self, "Save XLSX", default_name, "Excel Workbook (*.xlsx)")
        if not path:
            return

        headers_list: list[str] = []
        for c in visible_cols:
            h = model.headerData(c, Qt.Horizontal, Qt.DisplayRole) or ""
            headers_list.append(str(h).replace("\n", " • "))

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "ISTQB Applications"
            ws.append(headers_list)
            for r in range(model.rowCount()):
                row_vals = []
                for c in visible_cols:
                    val = model.index(r, c).data(Qt.DisplayRole)
                    row_vals.append("" if val is None else str(val))
                ws.append(row_vals)
            # Auto width
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(12, max_len + 2), 60)
            wb.save(path)
            QMessageBox.information(self, "Export XLSX", f"Exported {model.rowCount()} rows to:\n{path}")
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
    
        # >>> Přidáno: kontextové menu pro editaci v Overview
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.on_overview_context_menu)
        # <<<
    
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
            
    def on_overview_context_menu(self, pos) -> None:
        from PySide6.QtWidgets import QMenu
        idx = self.table.indexAt(pos)
        if not idx.isValid():
            return
        # Zajistíme výběr řádku pod kurzorem
        self.table.selectRow(idx.row())
        menu = QMenu(self.table)
        act_edit = menu.addAction("Edit…")
        chosen = menu.exec_(self.table.viewport().mapToGlobal(pos))
        if chosen == act_edit:
            self._edit_overview_row(idx.row())

    def _edit_overview_row(self, proxy_row: int) -> None:
        from PySide6.QtWidgets import (
            QDialog, QFormLayout, QLineEdit, QDialogButtonBox
        )
        from PySide6.QtGui import QIcon
        from PySide6.QtWidgets import QStyle
    
        model = self.table.model()
        if model is None:
            return
    
        # Najdi source model a source row
        if isinstance(model, QSortFilterProxyModel):
            sidx = model.mapToSource(model.index(proxy_row, 0))
            src = model.sourceModel()
            src_row = sidx.row()
        else:
            src = model
            src_row = proxy_row
    
        if src is None or src_row < 0:
            return
    
        cols = src.columnCount()
        file_col = cols - 1  # poslední sloupec = "File name" (needitovat)
    
        # Sloupce k editaci = aktuálně VIDITELNÍ v Overview (kromě File name)
        # (Eligibility 10..14 jsou v Overview skryté už v rescan(); pokud by byly odskryté, stanou se editovatelnými.)
        editable_cols: list[int] = []
        for c in range(cols):
            if c == file_col:
                continue
            # dotaz na viditelnost přes proxy (Overview tab)
            try:
                if not self.table.isColumnHidden(c):
                    editable_cols.append(c)
            except Exception:
                editable_cols.append(c)
    
        # Sestav dialog
        dlg = QDialog(self)
        dlg.setWindowTitle("Edit record")
        form = QFormLayout(dlg)
    
        editors: list[tuple[int, QLineEdit]] = []
        for c in editable_cols:
            header = src.headerData(c, Qt.Horizontal, Qt.DisplayRole) or f"Column {c}"
            header = str(header).replace("\n", " • ")
            val = src.index(src_row, c).data(Qt.DisplayRole)
            le = QLineEdit(dlg)
            le.setText("" if val is None else str(val))
            form.addRow(header + ":", le)
            editors.append((c, le))
    
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=dlg)
        form.addWidget(btns)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
    
        if dlg.exec_() != QDialog.Accepted:
            return
    
        # Zapis zpět do source modelu
        for c, le in editors:
            src.setData(src.index(src_row, c), le.text(), Qt.DisplayRole)
    
        # Refresh ikonek pro Wished Recognitions (Academia = 4, Certified = 5)
        def _set_yesno_icon(col: int) -> None:
            try:
                idx = src.index(src_row, col)
                text = (src.data(idx, Qt.DisplayRole) or "").strip().lower()
                icon_yes = self.style().standardIcon(QStyle.SP_DialogApplyButton)
                icon_no  = self.style().standardIcon(QStyle.SP_DialogCancelButton)
                # Nastavíme dekoraci položky (ekvivalent .setIcon u QStandardItem)
                src.setData(idx, icon_yes if text in {"yes","on","true","1","checked"} else icon_no, Qt.DecorationRole)
            except Exception:
                pass
    
        try:
            _set_yesno_icon(4)
            _set_yesno_icon(5)
        except Exception:
            pass
    
        # Info do status baru
        try:
            self.statusBar().showMessage("Overview: record edited.")
        except Exception:
            pass

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
        from pathlib import Path
        # --- 1) Robustní určení kořene PDF ---
        candidates = []
        if isinstance(getattr(self, "pdf_root", None), Path):
            candidates.append(self.pdf_root)
        try:
            candidates.append(Path(__file__).resolve().parent.parent / "PDF")
        except Exception:
            pass
        try:
            candidates.append(Path.cwd() / "PDF")
        except Exception:
            pass
    
        chosen = None
        for c in candidates:
            try:
                if isinstance(c, Path) and c.exists() and c.is_dir():
                    chosen = c
                    break
            except Exception:
                continue
        if chosen is None:
            chosen = candidates[0] if candidates else None
    
        if chosen is not None and chosen != getattr(self, "pdf_root", None):
            self.pdf_root = chosen
    
        # --- 2) Diagnostika počtu nalezených PDF (bez parsování) + IGNORE __archive__ ---
        try:
            found = 0
            if isinstance(self.pdf_root, Path):
                for p in self.pdf_root.rglob("*"):
                    if p.is_file() and p.suffix.lower() == ".pdf":
                        rel = p.relative_to(self.pdf_root)
                        if "__archive__" in rel.parts:
                            continue
                        found += 1
            else:
                found = 0
        except Exception:
            found = 0
    
        # --- 3) Skutečné parsování ---
        scanner = PdfScanner(self.pdf_root) if isinstance(self.pdf_root, Path) else None
        self.records = scanner.scan() if scanner else []
    
        # --- 4) Naplnění modelu ---
        from PySide6.QtGui import QStandardItemModel, QStandardItem, QBrush, QColor, QIcon
        from PySide6.QtWidgets import QStyle
    
        headers = [
            "Board",
            "Application\nApplication Type",
            "Name of Your Academic Institution\nInstitution Name",
            "Name of Your Academic Institution\nCandidate Name",
            "Wished Recognitions\nAcademia Recognition",
            "Wished Recognitions\nCertified Recognition",
            "Contact details for Information exchange\nFull Name",
            "Contact details for Information exchange\nEmail Address",
            "Contact details for Information exchange\nPhone Number",
            "Contact details for Information exchange\nPostal Address",
            "Eligibility Evidence\nSyllabi Integration",      # 10 (HIDE)
            "Eligibility Evidence\nCourses/Modules",          # 11 (HIDE)
            "Eligibility Evidence\nProof of ISTQB Certifications",  # 12 (HIDE)
            "Eligibility Evidence\nUniversity Links",         # 13 (HIDE)
            "Eligibility Evidence\nAdditional Info/Documents",# 14 (HIDE)
            "Signature\nSignature Date",
            "File\nFile name",
        ]
    
        model = QStandardItemModel(0, len(headers), self)
        model.setHorizontalHeaderLabels(headers)
    
        # Barevné skupiny
        COLS_APPLICATION = [1]
        COLS_INSTITUTION = [2, 3]
        COLS_RECOG      = [4, 5]
        COLS_CONTACT    = [6, 7, 8, 9]
        COLS_ELIG       = [10, 11, 12, 13, 14]
    
        BRUSH_APP   = QBrush(QColor(58, 74, 110))
        BRUSH_INST  = QBrush(QColor(74, 58, 110))
        BRUSH_RECOG = QBrush(QColor(58, 110, 82))
        BRUSH_CONT  = QBrush(QColor(110, 82, 58))
        BRUSH_ELIG  = QBrush(QColor(72, 110, 110))
    
        icon_yes = self.style().standardIcon(QStyle.SP_DialogApplyButton)
        icon_no  = self.style().standardIcon(QStyle.SP_DialogCancelButton)
    
        def paint_group(items: list[QStandardItem], cols: list[int], brush: QBrush) -> None:
            for c in cols:
                if 0 <= c < len(items):
                    items[c].setBackground(brush)
    
        def set_yesno_icon(item: QStandardItem) -> None:
            val = (item.text() or "").strip().lower()
            if val in {"yes", "on", "true", "1", "checked"}:
                item.setIcon(icon_yes)
            else:
                item.setIcon(icon_no)
    
        for rec in self.records:
            row_vals = rec.as_row()  # 17 prvků
            items = [QStandardItem(v) for v in row_vals]
            for it in items:
                it.setEditable(False)
    
            # Barvy skupin
            paint_group(items, COLS_APPLICATION, BRUSH_APP)
            paint_group(items, COLS_INSTITUTION, BRUSH_INST)
            paint_group(items, COLS_RECOG,      BRUSH_RECOG)
            paint_group(items, COLS_CONTACT,    BRUSH_CONT)
            paint_group(items, COLS_ELIG,       BRUSH_ELIG)
    
            # Ikony do sloupců "Wished Recognitions"
            set_yesno_icon(items[4])  # Academia
            set_yesno_icon(items[5])  # Certified
    
            # Skrytá plná cesta v posledním sloupci
            FILE_COL = len(headers) - 1
            items[FILE_COL].setData(str(rec.path), Qt.UserRole + 1)
            model.appendRow(items)
    
        proxy = RecordsModel(headers, self)
        proxy.setSourceModel(model)
        self.table.setModel(proxy)
    
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        for c in range(len(headers)):
            self.table.resizeColumnToContents(c)
    
        # Skryj vertikální číslování řádků (row header)
        try:
            self.table.verticalHeader().setVisible(False)
        except Exception:
            pass
    
        # Výchozí řazení: Board
        self.table.sortByColumn(0, Qt.AscendingOrder)
    
        # Skryj Eligibility sloupce v Overview
        for c in (10, 11, 12, 13, 14):
            self.table.setColumnHidden(c, True)
    
        # Stav – found/parsed/root (found bez __archive__)
        try:
            root_str = str(self.pdf_root) if isinstance(self.pdf_root, Path) else "<unset>"
            self.statusBar().showMessage(f"PDF found: {found} • Parsed: {len(self.records)} • Root: {root_str}")
        except Exception:
            pass
    
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

    def _renumber_rows(self) -> None:
        """Write 1..N into the 'No.' column in source model according to current proxy order."""
        proxy = self.table.model()
        if not isinstance(proxy, QSortFilterProxyModel):
            return
        src = proxy.sourceModel()
        if src is None:
            return
        # NUM col = 0
        for r in range(proxy.rowCount()):
            pidx = proxy.index(r, 0)  # any column maps row; col 0 exists
            sidx = proxy.mapToSource(pidx)
            # Set display text to 1-based row number
            src.setData(src.index(sidx.row(), 0), str(r + 1), Qt.DisplayRole)

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
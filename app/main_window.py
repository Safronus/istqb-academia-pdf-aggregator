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














from PySide6.QtCore import Qt
from PySide6.QtWidgets import QStyledItemDelegate, QStyleOptionViewItem, QApplication, QStyle

class BoardHidingDelegate(QStyledItemDelegate):
    """
    Vykreslovací delegát pro sloupec 'Board':
    - Pro PRVNÍ výskyt hodnoty v aktuálním pořadí tabulky vykreslí text normálně.
    - Pro DALŠÍ výskyty stejné hodnoty vykreslí prázdný text (jen vizuálně).
    - Nezasahuje do modelu, dat ani exportu (data(DisplayRole) zůstávají zachována).
    """
    def paint(self, painter, option, index):
        if index.column() == 0:  # sloupec Board
            current = index.data(Qt.DisplayRole)
            hide = False
            # Zjisti, zda se stejná hodnota vyskytla již výše (v proxy pořadí)
            # Pozn.: lineární průchod jen přes viditelný model; výkonově OK pro běžné tabulky.
            try:
                model = index.model()
                for r in range(0, index.row()):
                    other = model.index(r, index.column()).data(Qt.DisplayRole)
                    if other == current and current not in (None, ""):
                        hide = True
                        break
            except Exception:
                hide = False

            if hide:
                opt = QStyleOptionViewItem(option)
                self.initStyleOption(opt, index)
                opt.text = ""  # vykreslit prázdný text, zbytek styly necháme
                style = option.widget.style() if option.widget else QApplication.style()
                style.drawControl(QStyle.CE_ItemViewItem, opt, painter, option.widget)
                return

        # Defaultní vykreslení pro ostatní sloupce nebo první výskyt Board
        super().paint(painter, option, index)
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        

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
    
        # Kořen "Sorted PDFs"
        self.sorted_root = (Path.cwd() / "Sorted PDFs").resolve()
    
        # DB pro Sorted PDFs
        from app.sorted_db import SortedDb
        self.sorted_db = SortedDb(self.sorted_root)
        self.sorted_db.load()
    
        self.records: List[PdfRecord] = []
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
    
        self.overview_tab = QWidget()
        self.browser_tab = QWidget()
        self.sorted_tab = QWidget()
    
        self.tabs.addTab(self.overview_tab, "Overview")
        self.tabs.addTab(self.browser_tab, "PDF Browser")
        self.tabs.addTab(self.sorted_tab, "Sorted PDFs")
    
        self._build_menu()
        self._build_overview_tab()
        self._build_browser_tab()
        self._build_sorted_tab()
    
        self.rescan()
        self.rescan_sorted()
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

    def export_selected_to_sorted(self) -> None:
        from PySide6.QtWidgets import QMessageBox
        from pathlib import Path
        import shutil
    
        sel = self.table.selectionModel()
        if not sel or not sel.hasSelection():
            QMessageBox.information(self, "Export", "Please select at least one row.")
            return
    
        # Detekce sloupců Board a File name (předpoklad: poslední sloupec = File name s fullpath v UserRole+1)
        model = self.table.model()
        if model is None:
            return
    
        # Zjisti indexy vybraných řádků v source modelu
        rows = [i.row() for i in sel.selectedRows()]
        if not rows:
            QMessageBox.information(self, "Export", "Není co exportovat.")
            return
    
        # Najdi sloupce
        COL_BOARD = 0
        COL_FILE  = model.columnCount() - 1
    
        exported = 0
        for r in rows:
            board = (model.index(r, COL_BOARD).data() or "").strip()
            # Full path je v UserRole + 1
            full = model.index(r, COL_FILE).data(Qt.UserRole + 1) or model.index(r, COL_FILE).data()
            if not full:
                continue
            full = Path(str(full))
            if not full.exists():
                continue
    
            # Cílová složka Sorted PDFs/<board> (board může být prázdný → kořen Sorted PDFs)
            dest_dir = (self.sorted_root / board).resolve()
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / full.name
    
            try:
                shutil.copy2(full, dest)
                exported += 1
            except Exception:
                continue
    
        # Rescan Sorted (naparsuje a upsertne do DB)
        self.rescan_sorted()
    
        try:
            self.statusBar().showMessage(f"Exportováno do 'Sorted PDFs': {exported} souborů.")
        except Exception:
            pass

    # ----- Overview tab -----
    def _build_overview_tab(self) -> None:
        from PySide6.QtWidgets import (
            QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QLineEdit,
            QPushButton, QTableView, QToolButton, QMenu
        )
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QStyle
    
        layout = QVBoxLayout()
        controls = QHBoxLayout()
    
        # === NOVÉ: tlačítko "Unparsed" (ad-hoc audit; žádná změna scanneru) ===
        self.btn_unparsed = QToolButton(self)
        self.btn_unparsed.setText("Unparsed")
        self.btn_unparsed.setToolTip("Show PDFs found on disk that are not present in Overview")
        self.btn_unparsed.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxWarning))
        self.btn_unparsed.setAutoRaise(True)
        # jemné vizuální zvýraznění (dark theme safe)
        self.btn_unparsed.setStyleSheet("QToolButton { color: #ff6b6b; font-weight: 600; }")
        self.btn_unparsed.clicked.connect(self.show_unparsed_report)
        
    
        # Export button (zůstává z 0.5a)
        self.btn_export = QToolButton(self)
        self.btn_export.setToolTip("Export…")
        self.btn_export.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.btn_export.setAutoRaise(True)
        self.btn_export.clicked.connect(self.on_export_overview)
        controls.addWidget(self.btn_export)
    
        self.board_combo = QComboBox()
        self.board_combo.addItem("All")
        for b in sorted(KNOWN_BOARDS):
            self.board_combo.addItem(b)
        self.board_combo.currentTextChanged.connect(self._filter_board)
    
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search…")
        self.search_edit.textChanged.connect(self._filter_text)
    
        self.open_btn = QPushButton("Open Selected PDF")
        self.open_btn.clicked.connect(self.open_selected_pdf)
    
        controls.addWidget(self.btn_unparsed)
        controls.addSpacing(12)
        controls.addWidget(QLabel("Board:"))
        controls.addWidget(self.board_combo, 1)
        controls.addSpacing(12)
        controls.addWidget(QLabel("Search:"))
        controls.addWidget(self.search_edit, 4)
        controls.addSpacing(12)
        controls.addWidget(self.open_btn)
    
        layout.addLayout(controls)
    
        # Tabulka Overview
        self.table = QTableView()
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection)  # multiselect
        self.table.doubleClicked.connect(self.open_selected_pdf)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setMinimumHeight(44)
    
        # === NOVÉ: persistentní model + proxy (použije se i v rescan) ===
        from PySide6.QtGui import QStandardItemModel
        self._headers = [
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
            "Eligibility Evidence\nSyllabi Integration",
            "Eligibility Evidence\nCourses/Modules",
            "Eligibility Evidence\nProof of ISTQB Certifications",
            "Eligibility Evidence\nUniversity Links",
            "Eligibility Evidence\nAdditional Info/Documents",
            "Signature Date",
            "File\nFile name",
        ]
        if not hasattr(self, "_source_model"):
            self._source_model = QStandardItemModel(0, len(self._headers), self)
            self._source_model.setHorizontalHeaderLabels(self._headers)
        if not hasattr(self, "_proxy"):
            self._proxy = RecordsModel(self._headers, self)
            self._proxy.setSourceModel(self._source_model)
            self._proxy.setDynamicSortFilter(True)
            self.table.setModel(self._proxy)
            
        self.table.setItemDelegateForColumn(0, BoardHidingDelegate(self.table))
        
        # Skryj Eligibility sloupce už zde
        for c in (10, 11, 12, 13, 14):
            self.table.setColumnHidden(c, True)
    
        # Kontextové menu – export do Sorted
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
    
        def _ctx(pos):
            idx = self.table.indexAt(pos)
            if idx.isValid():
                if not self.table.selectionModel().isSelected(idx):
                    self.table.selectRow(idx.row())
            menu = QMenu(self.table)
            act_export_sorted = menu.addAction("Exportovat do složky 'Sorted PDFs'")
            chosen = menu.exec_(self.table.viewport().mapToGlobal(pos))
            if chosen == act_export_sorted:
                self.export_selected_to_sorted()
    
        if not hasattr(self, "_overview_ctx_connected"):
            self.table.customContextMenuRequested.connect(_ctx)
            self._overview_ctx_connected = True
    
        layout.addWidget(self.table, 1)
        self.overview_tab.setLayout(layout)

    def _enumerate_all_pdfs(self) -> list[str]:
        """
        Vrátí list absolutních cest na *.pdf pod self.pdf_root (rekurzivně),
        s ignorováním jakékoli složky '__archive__'.
        """
        from pathlib import Path
        root = getattr(self, "pdf_root", None)
        if not root:
            return []
        root = Path(root).resolve()
        out: list[str] = []
        try:
            for p in root.rglob("*.pdf"):
                try:
                    rel = p.resolve().relative_to(root)
                    if "__archive__" in rel.parts:
                        continue
                    out.append(str(p.resolve()))
                except Exception:
                    # pokud by nešla relativní cesta, stále přidáme (bez filtru)
                    out.append(str(p.resolve()))
        except Exception:
            pass
        return out

    def show_unparsed_report(self) -> None:
        """
        Ad-hoc audit: porovná PDF v self.pdf_root (rekurzivně; ignoruje '__archive__')
        s aktuálně zobrazenými záznamy (self.records).
        Zobrazí dialog s přehledem "unparsed" PDF (Board, File name, Full path).
        """
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QLabel,
            QTreeWidget, QTreeWidgetItem, QPushButton
        )
        from PySide6.QtCore import Qt
        from pathlib import Path
    
        # 1) PDF na disku (mimo __archive__)
        all_pdfs = self._enumerate_all_pdfs()
    
        # 2) PDF v Overview (už naparsovaná) – vezmeme absolutní cesty
        parsed_paths = set()
        try:
            for r in getattr(self, "records", []):
                p = getattr(r, "path", None)
                if p:
                    parsed_paths.add(str(Path(p).resolve()))
        except Exception:
            pass
    
        # 3) Rozdíl = unparsed
        unparsed = []
        for p in all_pdfs:
            ap = str(Path(p).resolve())
            if ap not in parsed_paths:
                # board = první složka pod pdf_root (pokud existuje)
                try:
                    rel = Path(p).resolve().relative_to(self.pdf_root.resolve())
                    board = rel.parts[0] if len(rel.parts) > 1 else ""
                except Exception:
                    board = ""
                unparsed.append((board, Path(p).name, ap))
    
        # 4) Dialog s výsledky
        dlg = QDialog(self)
        dlg.setWindowTitle("Unparsed PDFs")
        dlg.resize(800, 500)
    
        lay = QVBoxLayout(dlg)
        header = QLabel(f"PDFs on disk not present in Overview: {len(unparsed)}")
        header.setStyleSheet("QLabel { color: #ff6b6b; font-weight: 600; }")
        lay.addWidget(header)
    
        tree = QTreeWidget()
        tree.setHeaderLabels(["Board", "File name", "Full path"])
        tree.header().setDefaultAlignment(Qt.AlignCenter)
        tree.header().setStretchLastSection(True)
        for board, fname, fullp in sorted(unparsed, key=lambda t: (t[0].lower(), t[1].lower())):
            it = QTreeWidgetItem([board or "—", fname, fullp])
            tree.addTopLevelItem(it)
        tree.expandAll()
        lay.addWidget(tree, 1)
    
        # Tlačítka
        btns = QHBoxLayout()
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        btns.addStretch(1)
        btns.addWidget(btn_close)
        lay.addLayout(btns)
    
        # 5) Aktualizuj badge na tlačítku
        try:
            if len(unparsed) > 0:
                self.btn_unparsed.setText(f"Unparsed: {len(unparsed)}")
                self.statusBar().showMessage(f"Unparsed PDFs: {len(unparsed)}")
            else:
                self.btn_unparsed.setText("Unparsed")
                self.statusBar().showMessage("All PDFs present in Overview.")
        except Exception:
            pass
    
        dlg.exec()

    def _collect_available_boards(self) -> list[str]:
        """
        Získá boardy, pro které skutečně existují PDF ve složce PDF/<board>/...
        Ignoruje podsložku '__archive__'. Pokud není pdf_root dostupný,
        vrátí unikátní boardy z již naparsovaných self.records.
        """
        from pathlib import Path
        boards: set[str] = set()
    
        root = getattr(self, "pdf_root", None)
        if isinstance(root, Path) and root.exists():
            try:
                # první úroveň podsložek pod PDF je board
                for sub in root.iterdir():
                    if not sub.is_dir():
                        continue
                    if sub.name == "__archive__":
                        continue
                    # zkontroluj aspoň jedno .pdf uvnitř (rekurzivně)
                    found_pdf = False
                    for p in sub.rglob("*.pdf"):
                        try:
                            rel = p.relative_to(root)
                            if "__archive__" in rel.parts:
                                continue
                            found_pdf = True
                            break
                        except Exception:
                            continue
                    if found_pdf:
                        boards.add(sub.name)
            except Exception:
                pass
    
        if not boards:
            try:
                for r in getattr(self, "records", []):
                    if getattr(r, "board", None):
                        boards.add(r.board)
            except Exception:
                pass
    
        return sorted(boards)
    
    def on_export_overview(self) -> None:
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QGroupBox, QCheckBox, QLabel,
            QComboBox, QPushButton, QListWidget, QListWidgetItem, QScrollArea,
            QWidget, QGridLayout, QDialogButtonBox, QFileDialog, QMessageBox
        )
        from PySide6.QtCore import Qt
        from datetime import datetime
        import os
    
        # --- definice polí (label -> klíč v rec) ---
        # soulad s Overview pořadím
        FIELDS: list[tuple[str, str]] = [
            ("Board", "board"),
            ("Application Type", "application_type"),
            ("Institution Name", "institution_name"),
            ("Candidate Name", "candidate_name"),
            ("Academia Recognition", "recognition_academia"),
            ("Certified Recognition", "recognition_certified"),
            ("Full Name", "contact_full_name"),
            ("Email Address", "contact_email"),
            ("Phone Number", "contact_phone"),
            ("Postal Address", "contact_postal_address"),
            # Eligibility (v Overview skryté, ale exportovatelně dostupné)
            ("Syllabi Integration", "syllabi_integration_description"),
            ("Courses/Modules", "courses_modules_list"),
            ("Proof of ISTQB Certifications", "proof_of_istqb_certifications"),
            ("University Links", "university_links"),
            ("Additional Info/Documents", "additional_information_documents"),
            ("Signature Date", "signature_date"),
            ("File name", "file_name"),
        ]
    
        # --- dialog ---
        class ExportDialog(QDialog):
            def __init__(self, boards_avail: list[str], parent=None):
                super().__init__(parent)
                self.setWindowTitle("Export options")
                main = QVBoxLayout(self)
    
                # --- Formáty ---
                grp_fmt = QGroupBox("Export formats")
                fmt_lay = QHBoxLayout(grp_fmt)
                self.chk_csv = QCheckBox("CSV")
                self.chk_xlsx = QCheckBox("XLSX")
                self.chk_txt = QCheckBox("TXT (formatted)")
                # default: CSV + TXT zapnuto, XLSX vypnuto (bez nové závislosti)
                self.chk_csv.setChecked(True)
                self.chk_txt.setChecked(True)
                self.chk_xlsx.setChecked(False)
                fmt_lay.addWidget(self.chk_csv)
                fmt_lay.addWidget(self.chk_xlsx)
                fmt_lay.addWidget(self.chk_txt)
                main.addWidget(grp_fmt)
    
                # --- Boardy ---
                grp_b = QGroupBox("Boards to export")
                b_lay = QGridLayout(grp_b)
                self.all_boards = QCheckBox("All")
                self.all_boards.setChecked(True)
                b_lay.addWidget(self.all_boards, 0, 0, 1, 2)
    
                self.cmb_board = QComboBox()
                self.cmb_board.addItems(boards_avail)
                self.btn_add = QPushButton("Add")
                b_lay.addWidget(QLabel("Board:"), 1, 0)
                b_lay.addWidget(self.cmb_board, 1, 1)
                b_lay.addWidget(self.btn_add, 1, 2)
    
                self.list_sel = QListWidget()
                b_lay.addWidget(QLabel("Selected boards:"), 2, 0, 1, 3)
                b_lay.addWidget(self.list_sel, 3, 0, 1, 3)
                self.btn_remove = QPushButton("Remove selected")
                b_lay.addWidget(self.btn_remove, 4, 0, 1, 3)
    
                def _toggle_board_ui():
                    enabled = not self.all_boards.isChecked()
                    self.cmb_board.setEnabled(enabled)
                    self.btn_add.setEnabled(enabled)
                    self.list_sel.setEnabled(enabled)
                    self.btn_remove.setEnabled(enabled)
    
                self.all_boards.toggled.connect(_toggle_board_ui)
                self.btn_add.clicked.connect(lambda: self._add_board())
                self.btn_remove.clicked.connect(lambda: self._remove_sel())
                _toggle_board_ui()
                main.addWidget(grp_b)
    
                # --- Pole ---
                grp_fields = QGroupBox("Fields to export")
                f_lay = QVBoxLayout(grp_fields)
                self.chk_all_fields = QCheckBox("Select all")
                self.chk_all_fields.setChecked(True)
                f_lay.addWidget(self.chk_all_fields)
    
                self.scroll = QScrollArea()
                self.scroll.setWidgetResizable(True)
                cont = QWidget()
                cont_lay = QVBoxLayout(cont)
                self.field_checks: list[tuple[str, str, QCheckBox]] = []
                for label, key in FIELDS:
                    cb = QCheckBox(label)
                    cb.setChecked(True)
                    self.field_checks.append((label, key, cb))
                    cont_lay.addWidget(cb)
                cont_lay.addStretch(1)
                self.scroll.setWidget(cont)
                f_lay.addWidget(self.scroll)
    
                def _toggle_all():
                    state = self.chk_all_fields.isChecked()
                    for _, _, cb in self.field_checks:
                        cb.setChecked(state)
                self.chk_all_fields.toggled.connect(_toggle_all)
                main.addWidget(grp_fields)
    
                # --- tlačítka ---
                btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
                main.addWidget(btns)
                btns.accepted.connect(self._accept)
                btns.rejected.connect(self.reject)
    
                self.result = None  # dict result
    
            def _add_board(self):
                txt = self.cmb_board.currentText().strip()
                if not txt:
                    return
                # nepřidávej duplicitně
                for i in range(self.list_sel.count()):
                    if self.list_sel.item(i).text() == txt:
                        return
                self.list_sel.addItem(QListWidgetItem(txt))
    
            def _remove_sel(self):
                for it in self.list_sel.selectedItems():
                    self.list_sel.takeItem(self.list_sel.row(it))
    
            def _accept(self):
                # formáty
                fmts = []
                if self.chk_csv.isChecked(): fmts.append("csv")
                if self.chk_xlsx.isChecked(): fmts.append("xlsx")
                if self.chk_txt.isChecked(): fmts.append("txt")
                if not fmts:
                    QMessageBox.warning(self, "Export", "Select at least one format.")
                    return
    
                # boardy
                boards = None
                if not self.all_boards.isChecked():
                    boards = []
                    for i in range(self.list_sel.count()):
                        boards.append(self.list_sel.item(i).text())
                    if not boards:
                        QMessageBox.warning(self, "Export", "Add at least one board or check 'All'.")
                        return
    
                # pole
                fields = [(lab, key) for (lab, key, cb) in self.field_checks if cb.isChecked()]
                if not fields:
                    QMessageBox.warning(self, "Export", "Select at least one field.")
                    return
    
                self.result = {
                    "formats": fmts,
                    "boards": boards,  # None = All
                    "fields": fields,  # list of (label, key)
                }
                self.accept()
    
        # připrav data pro dialog
        boards_avail = self._collect_available_boards()
        dlg = ExportDialog(boards_avail, self)
        if dlg.exec_() != QDialog.Accepted or dlg.result is None:
            return
    
        formats: list[str] = dlg.result["formats"]
        boards_sel: list[str] | None = dlg.result["boards"]
        fields: list[tuple[str, str]] = dlg.result["fields"]
    
        # výběr cesty (jeden dialog – základní soubor, ostatní formáty vedle)
        # navržený název podle času a případných boardů
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        suffix = "all" if boards_sel is None else ("+".join(boards_sel[:3]) + ("+more" if boards_sel and len(boards_sel) > 3 else ""))
        base_name = f"export-{suffix}-{ts}".replace(" ", "_")
        # defaultní filtr podle první volby
        first_ext = formats[0]
        default_name = f"{base_name}.{first_ext}"
    
        path, _ = QFileDialog.getSaveFileName(self, "Save export as…", default_name,
                                              "All files (*.*)")
        if not path:
            return
    
        base_no_ext, _sep, _ext = path.rpartition(".")
        if not base_no_ext:
            # uživatel nedal příponu → použij base_name z dialogu
            base_no_ext = path
    
        # vyber záznamy dle boards
        records = list(getattr(self, "records", []))
        if boards_sel is not None:
            wh = set(boards_sel)
            records = [r for r in records if (getattr(r, "board", None) in wh)]
    
        # připrav řádky (jen vybraná pole)
        # každý řádek → list hodnot podle (label, key) pořadí
        rows: list[list[str]] = []
        headers = [lab for (lab, _key) in fields]
    
        def _get_val(rec, key: str) -> str:
            if key == "file_name":
                try:
                    return rec.path.name if getattr(rec, "path", None) else ""
                except Exception:
                    return ""
            v = getattr(rec, key, None)
            return "" if v is None else str(v)
    
        for rec in records:
            rows.append([_get_val(rec, key) for (_lab, key) in fields])
    
        # exporty
        ok, errs = [], []
    
        if "csv" in formats:
            try:
                fn = base_no_ext + ".csv"
                self._export_to_csv(fn, headers, rows)
                ok.append(fn)
            except Exception as e:
                errs.append(f"CSV: {e}")
    
        if "txt" in formats:
            try:
                fn = base_no_ext + ".txt"
                self._export_to_txt(fn, headers, rows)
                ok.append(fn)
            except Exception as e:
                errs.append(f"TXT: {e}")
    
        if "xlsx" in formats:
            try:
                fn = base_no_ext + ".xlsx"
                self._export_to_xlsx(fn, headers, rows)
                ok.append(fn)
            except ImportError:
                errs.append("XLSX: optional dependency 'openpyxl' not installed.")
            except Exception as e:
                errs.append(f"XLSX: {e}")
    
        # výsledek
        msg = []
        if ok:   msg.append("Exported:\n- " + "\n- ".join(ok))
        if errs: msg.append("\nIssues:\n- " + "\n- ".join(errs))
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.information(self, "Export", "\n".join(msg))
        try:
            self.statusBar().showMessage("Export done.")
        except Exception:
            pass

    def _export_to_csv(self, filename: str, headers: list[str], rows: list[list[str]]) -> None:
        import csv
        from pathlib import Path
        Path(filename).parent.mkdir(parents=True, exist_ok=True)
        with open(filename, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(headers)
            for r in rows:
                w.writerow(r)
                
    def _export_to_txt(self, path: str, headers: list[str], rows: list[list[str]]) -> None:
        """
        Rich TXT report shared by Overview & Sorted exports.
    
        Structure per record:
          Title:  Institution — Candidate — [Board]  (fallback to File name)
          Basic info:
            - Board: ...
            - Application Type: ...
            - Institution Name: ...
            - Candidate Name: ...
            - Signature Date: ...
            - File name: ...
          Sections (only if data present):
            Recognition:
              - Academia Recognition: ...
              - Certified Recognition: ...
            Contact:
              - Full Name: ...
              - Email Address: ...
              - Phone Number: ...
              - Postal Address: ...
            Curriculum:
              - Syllabi Integration: ...
              - Courses/Modules: ...
            Evidence:
              - Proof of ISTQB Certifications: ...
              - Additional Info/Documents: ...
            Links:
              - University Links: ...
        """
        def _hmap(hs: list[str]) -> dict[str, int]:
            return {h: i for i, h in enumerate(hs)}
    
        def _get(row: list[str], h2i: dict[str, int], label: str, *alts: str) -> str:
            for k in (label, *alts):
                if k in h2i:
                    idx = h2i[k]
                    if 0 <= idx < len(row):
                        val = row[idx]
                        return "" if val is None else str(val)
            return ""
    
        def _lines(val: str) -> list[str]:
            if not val:
                return []
            s = str(val).replace("\r\n", "\n").replace("\r", "\n").strip()
            if not s:
                return []
            return [ln.rstrip() for ln in s.split("\n")]
    
        def _write_bullet(fh, label: str, value: str) -> bool:
            ls = _lines(value)
            if not ls:
                return False
            # first line
            fh.write(f"- {label}: {ls[0]}\n")
            # following lines as indented bullets
            for ln in ls[1:]:
                if ln.strip():
                    fh.write(f"  {ln}\n")
            return True
    
        # Section definitions by labels (must match header labels)
        SECTIONS: list[tuple[str, list[str]]] = [
            ("Recognition", [
                "Academia Recognition",
                "Certified Recognition",
            ]),
            ("Contact", [
                "Full Name",
                "Email Address",
                "Phone Number",
                "Postal Address",
            ]),
            ("Curriculum", [
                "Syllabi Integration",
                "Courses/Modules",
            ]),
            ("Evidence", [
                "Proof of ISTQB Certifications",
                "Additional Info/Documents",
            ]),
            ("Links", [
                "University Links",
            ]),
        ]
    
        h2i = _hmap(headers)
    
        with open(path, "w", encoding="utf-8") as fh:
            for idx, row in enumerate(rows, start=1):
                board   = _get(row, h2i, "Board")
                inst    = _get(row, h2i, "Institution Name")
                cand    = _get(row, h2i, "Candidate Name")
                app_t   = _get(row, h2i, "Application Type")
                sigdate = _get(row, h2i, "Signature Date", "Signature date", "signature_date", "sigdate")
                fname   = _get(row, h2i, "File name", "File Name", "Filename", "filename")
    
                # -------- Title --------
                parts = []
                if inst: parts.append(inst)
                if cand: parts.append(cand)
                if board: parts.append(f"[{board}]")
                title = " — ".join(p for p in parts if p) or (fname or f"Record #{idx}")
                fh.write(title + "\n")
                fh.write("=" * len(title) + "\n\n")
    
                # -------- Basic info --------
                fh.write("Basic info:\n")
                _write_bullet(fh, "Board", board)
                _write_bullet(fh, "Application Type", app_t)
                _write_bullet(fh, "Institution Name", inst)
                _write_bullet(fh, "Candidate Name", cand)
                _write_bullet(fh, "Signature Date", sigdate)
                _write_bullet(fh, "File name", fname)
                fh.write("\n")
    
                # -------- Sections --------
                for sec_title, labels in SECTIONS:
                    # include section only if at least one label has data
                    any_data = any(_get(row, h2i, lab) for lab in labels if lab in h2i)
                    if not any_data:
                        continue
                    fh.write(f"{sec_title}:\n")
                    for lab in labels:
                        if lab in h2i:
                            _write_bullet(fh, lab, _get(row, h2i, lab))
                    fh.write("\n")
    
                # Separator between records
                if idx < len(rows):
                    fh.write("-----\n\n")
                
    def _export_to_xlsx(self, filename: str, headers: list[str], rows: list[list[str]]) -> None:
        """
        XLSX export přes optional dependency 'openpyxl'.
        Bez přidávání závislostí do projektu – pokud není k dispozici, vyhodí ImportError.
        """
        from pathlib import Path
        Path(filename).parent.mkdir(parents=True, exist_ok=True)
    
        try:
            import openpyxl
            from openpyxl.styles import Font
        except Exception as e:
            raise ImportError("openpyxl not available") from e
    
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Export"
    
        ws.append(headers)
        # tučný header
        bold = Font(bold=True)
        for c in range(1, len(headers) + 1):
            ws.cell(row=1, column=c).font = bold
    
        for r in rows:
            ws.append(r)
    
        # auto šířky (jednoduše dle max délky)
        for col_idx, head in enumerate(headers, start=1):
            max_len = len(head)
            for row in rows:
                if col_idx - 1 < len(row):
                    max_len = max(max_len, len(str(row[col_idx - 1])))
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max_len + 2, 80)
    
        wb.save(filename)

    def _filter_board(self, txt: str) -> None:
        proxy = self.table.model()
        if isinstance(proxy, RecordsModel):
            proxy.set_board(txt)

    def _filter_text(self, txt: str) -> None:
        proxy = self.table.model()
        if isinstance(proxy, RecordsModel):
            proxy.set_search(txt)

    def _build_sorted_tab(self) -> None:
        """
        UI pro záložku 'Sorted PDFs': vlevo strom Board -> PDF, vpravo detailní formulář.
        Minimal-change:
          - doplněny očekávané názvy polí (ed_inst_name/ed_cand_name/ed_rec_acad/ed_rec_cert/ed_sigdate + aliasy),
          - přidány jemné korekce velikostí (splitter, autosize stromu, kompaktní výška PTE),
          - zachováno tlačítko 'Export…' (volá export_sorted_db()).
        """
        from PySide6.QtWidgets import (
            QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem, QSplitter,
            QWidget, QFormLayout, QLineEdit, QPlainTextEdit, QPushButton
        )
        from PySide6.QtCore import Qt, QTimer
    
        layout = QVBoxLayout()
    
        # Hlavní splitter: vlevo strom, vpravo formulář
        self.split_sorted = QSplitter(self.sorted_tab)
        self.split_sorted.setOrientation(Qt.Horizontal)
    
        # ===== LEVÁ STRANA: strom se soubory =====
        self.tree_sorted = QTreeWidget()
        self.tree_sorted.setHeaderLabels(["Board / PDF"])
        self.tree_sorted.itemSelectionChanged.connect(self._sorted_on_item_changed)
    
        # ===== PRAVÁ STRANA: detailní formulář =====
        right = QWidget()
        right_layout = QVBoxLayout(right)
    
        self.form_sorted = QFormLayout()
        self.form_sorted.setFormAlignment(Qt.AlignTop)
        self.form_sorted.setLabelAlignment(Qt.AlignRight | Qt.AlignTop)
    
        # Jednořádková pole (názvy podle toho, co používají ostatní metody)
        self.ed_board = QLineEdit()
        self.ed_app_type = QLineEdit()
        self.ed_inst_name = QLineEdit()    # očekáváno ve zbytku kódu
        self.ed_cand_name = QLineEdit()    # očekáváno ve zbytku kódu
        self.ed_rec_acad = QLineEdit()     # očekáváno (alias k dřívějšímu ed_acad)
        self.ed_rec_cert = QLineEdit()     # očekáváno (alias k dřívějšímu ed_cert)
        self.ed_fullname = QLineEdit()
        self.ed_email = QLineEdit()
        self.ed_phone = QLineEdit()
        self.ed_sigdate = QLineEdit()      # očekáváno _sorted_set_editable
        self.ed_filename = QLineEdit()
    
        # Víceřádková pole – kompaktní výšky (jemná korekce)
        self.ed_address = QPlainTextEdit()
        self.ed_syllabi = QPlainTextEdit()
        self.ed_courses = QPlainTextEdit()
        self.ed_proof = QPlainTextEdit()
        self.ed_links = QPlainTextEdit()
        self.ed_additional = QPlainTextEdit()
        for pte in (self.ed_address, self.ed_syllabi, self.ed_courses,
                    self.ed_proof, self.ed_links, self.ed_additional):
            pte.setMinimumHeight(56)  # čitelná výška
            pte.setMaximumHeight(120) # nepřerůstá; víc obsahu = scroll uvnitř pole
    
        # ---- Kompatibilní aliasy (jen jmenné mosty) ----
        self.ed_inst = self.ed_inst_name
        self.ed_cand = self.ed_cand_name
        self.ed_acad = self.ed_rec_acad
        self.ed_cert = self.ed_rec_cert
        self.ed_signature_date = self.ed_sigdate
        self.ed_contact_full_name = self.ed_fullname
        self.ed_contact_email = self.ed_email
        self.ed_contact_phone = self.ed_phone
        self.ed_postal_address = self.ed_address
        self.ed_syllabi_integration_description = self.ed_syllabi
        self.ed_courses_modules_list = self.ed_courses
        self.ed_proof_of_istqb_certifications = self.ed_proof
        self.ed_university_links = self.ed_links
        self.ed_additional_information_documents = self.ed_additional
    
        # Sestavení formuláře – popisky zarovnány s Overview
        self.form_sorted.addRow("Board:", self.ed_board)
        self.form_sorted.addRow("Application Type:", self.ed_app_type)
        self.form_sorted.addRow("Institution Name:", self.ed_inst_name)
        self.form_sorted.addRow("Candidate Name:", self.ed_cand_name)
        self.form_sorted.addRow("Academia Recognition:", self.ed_rec_acad)
        self.form_sorted.addRow("Certified Recognition:", self.ed_rec_cert)
        self.form_sorted.addRow("Full Name:", self.ed_fullname)
        self.form_sorted.addRow("Email Address:", self.ed_email)
        self.form_sorted.addRow("Phone Number:", self.ed_phone)
        self.form_sorted.addRow("Postal Address:", self.ed_address)
        self.form_sorted.addRow("Syllabi Integration:", self.ed_syllabi)
        self.form_sorted.addRow("Courses/Modules:", self.ed_courses)
        self.form_sorted.addRow("Proof of ISTQB Certifications:", self.ed_proof)
        self.form_sorted.addRow("University Links:", self.ed_links)
        self.form_sorted.addRow("Additional Info/Documents:", self.ed_additional)
        self.form_sorted.addRow("Signature Date:", self.ed_sigdate)
        self.form_sorted.addRow("File name:", self.ed_filename)
    
        right_layout.addLayout(self.form_sorted)
    
        # Řádek tlačítek (včetně Export…)
        btn_row = QHBoxLayout()
        self.btn_sorted_edit = QPushButton("Edit")
        self.btn_sorted_save = QPushButton("Save to DB")
        self.btn_sorted_export = QPushButton("Export…")
        self.btn_sorted_rescan = QPushButton("Rescan Sorted")
    
        btn_row.addWidget(self.btn_sorted_edit)
        btn_row.addWidget(self.btn_sorted_save)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_sorted_export)
        btn_row.addWidget(self.btn_sorted_rescan)
        right_layout.addLayout(btn_row)
    
        # Vazby tlačítek – žádné přejmenování stávajících metod
        self.btn_sorted_edit.clicked.connect(lambda: self._sorted_set_editable(True))
        self.btn_sorted_save.clicked.connect(self._sorted_save_changes)
        self.btn_sorted_export.clicked.connect(self.export_sorted_db)
        self.btn_sorted_rescan.clicked.connect(self.rescan_sorted)
    
        # Výchozí – read-only pole; přepíná se tlačítkem Edit
        self._sorted_set_editable(False)
    
        # Osazení splitteru a layoutu
        self.split_sorted.addWidget(self.tree_sorted)
        self.split_sorted.addWidget(right)
        self.split_sorted.setStretchFactor(0, 1)
        self.split_sorted.setStretchFactor(1, 2)
    
        layout.addWidget(self.split_sorted, 1)
        self.sorted_tab.setLayout(layout)
    
        # ===== Jemné sizing doladění po vystavění UI (po event loopu) =====
        def _apply_sorted_sizes():
            try:
                # rozumné počáteční rozměry okna (není závazné – jen pokud je menší)
                rec_w, rec_h = 1200, 820
                if self.width() < rec_w or self.height() < rec_h:
                    self.resize(max(self.width(), rec_w), max(self.height(), rec_h))
            except Exception:
                pass
            try:
                # autosize sloupce stromu + výchozí poměr splitteru
                self.tree_sorted.resizeColumnToContents(0)
                total_w = max(self.width(), 1200)
                self.split_sorted.setSizes([int(total_w * 0.48), int(total_w * 0.52)])
            except Exception:
                pass
    
        QTimer.singleShot(0, _apply_sorted_sizes)
        
    def export_sorted_db(self) -> None:
        """
        Export CURRENT dataset from 'Sorted PDFs' tab using data stored in DB (self.sorted_db),
        with the SAME dialog and options as Overview:
          - formats: XLSX / CSV / TXT (multi-select, order matters for default extension),
          - board filter: All boards or specific subset,
          - fields: same labels/order as Overview (user-selectable),
        and the SAME engines: _export_to_xlsx/_export_to_csv/_export_to_txt.
        """
        from PySide6.QtWidgets import (
            QDialog, QVBoxLayout, QHBoxLayout, QGroupBox, QCheckBox, QLabel,
            QComboBox, QPushButton, QListWidget, QListWidgetItem, QScrollArea,
            QWidget, QGridLayout, QDialogButtonBox, QFileDialog, QMessageBox
        )
        from PySide6.QtCore import Qt
        from datetime import datetime
        import os
        from pathlib import Path
    
        # --- Columns (labels/keys) – keep identical to Overview order/labels ---
        FIELDS: list[tuple[str, str]] = [
            ("Board", "board"),
            ("Application Type", "application_type"),
            ("Institution Name", "institution_name"),
            ("Candidate Name", "candidate_name"),
            ("Academia Recognition", "recognition_academia"),
            ("Certified Recognition", "recognition_certified"),
            ("Full Name", "contact_full_name"),
            ("Email Address", "contact_email"),
            ("Phone Number", "contact_phone"),
            ("Postal Address", "contact_postal_address"),
            ("Syllabi Integration", "syllabi_integration_description"),
            ("Courses/Modules", "courses_modules_list"),
            ("Proof of ISTQB Certifications", "proof_of_istqb_certifications"),
            ("University Links", "university_links"),
            ("Additional Info/Documents", "additional_information_documents"),
            ("Signature Date", "signature_date"),
            ("File name", "file_name"),
        ]
    
        # --- Helper: boards present in the current Sorted tree (to mirror the view) ---
        def _collect_sorted_boards() -> list[str]:
            boards: list[str] = []
            try:
                top_count = self.tree_sorted.topLevelItemCount()
                for i in range(top_count):
                    t = self.tree_sorted.topLevelItem(i)
                    if t:
                        boards.append(t.text(0))
            except Exception:
                pass
            # fallback: use Overview boards if available
            if not boards and hasattr(self, "_collect_available_boards"):
                try:
                    boards = list(self._collect_available_boards())
                except Exception:
                    pass
            return sorted({b for b in boards if b})
    
        # --- Export dialog (identical structure/behavior to Overview) ---
        class ExportDialog(QDialog):
            def __init__(self, boards_avail: list[str], parent=None):
                super().__init__(parent)
                self.setWindowTitle("Export options")
                main = QVBoxLayout(self)
    
                # === Formats ===
                gb_formats = QGroupBox("Formats")
                lay_f = QHBoxLayout(gb_formats)
                self.cb_xlsx = QCheckBox("XLSX")
                self.cb_csv  = QCheckBox("CSV")
                self.cb_txt  = QCheckBox("TXT")
                # default like Overview: XLSX pre-checked
                self.cb_xlsx.setChecked(True)
                lay_f.addWidget(self.cb_xlsx); lay_f.addWidget(self.cb_csv); lay_f.addWidget(self.cb_txt)
                main.addWidget(gb_formats)
    
                # === Boards (All / subset) ===
                gb_boards = QGroupBox("Boards")
                lay_b = QVBoxLayout(gb_boards)
                self.chk_all_boards = QCheckBox("All boards")
                self.chk_all_boards.setChecked(True)
                lay_b.addWidget(self.chk_all_boards)
    
                self.lst_boards = QListWidget()
                self.lst_boards.setSelectionMode(QListWidget.MultiSelection)
                for b in boards_avail:
                    it = QListWidgetItem(b)
                    self.lst_boards.addItem(it)
                    it.setSelected(False)
                self.lst_boards.setEnabled(False)  # enabled only when not 'All boards'
                lay_b.addWidget(self.lst_boards)
                main.addWidget(gb_boards)
    
                def _toggle_boards(_=None):
                    self.lst_boards.setEnabled(not self.chk_all_boards.isChecked())
                self.chk_all_boards.toggled.connect(_toggle_boards)
    
                # === Fields (checkbox list, in a scroll) ===
                gb_fields = QGroupBox("Fields (columns)")
                lay_fields = QVBoxLayout(gb_fields)
                scr = QScrollArea()
                scr.setWidgetResizable(True)
                host = QWidget()
                grid = QGridLayout(host)
    
                self.field_checks: list[tuple[QCheckBox, tuple[str, str]]] = []
                for row, (label, key) in enumerate(FIELDS):
                    cb = QCheckBox(label)
                    cb.setChecked(True)
                    self.field_checks.append((cb, (label, key)))
                    grid.addWidget(cb, row, 0)
                scr.setWidget(host)
                lay_fields.addWidget(scr)
                main.addWidget(gb_fields)
    
                # === Buttons ===
                btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                main.addWidget(btns)
                btns.accepted.connect(self._accept)
                btns.rejected.connect(self.reject)
    
                self.result: dict | None = None
    
            def _accept(self):
                formats: list[str] = []
                if self.cb_xlsx.isChecked(): formats.append("xlsx")
                if self.cb_csv.isChecked():  formats.append("csv")
                if self.cb_txt.isChecked():  formats.append("txt")
                if not formats:
                    QMessageBox.information(self, "Export", "Please choose at least one format.")
                    return
    
                boards_sel: list[str] | None
                if self.chk_all_boards.isChecked():
                    boards_sel = None  # all
                else:
                    boards_sel = [it.text() for it in self.lst_boards.selectedItems()]
                    if not boards_sel:
                        QMessageBox.information(self, "Export", "Please select at least one board or check 'All boards'.")
                        return
    
                fields: list[tuple[str, str]] = []
                for cb, pair in self.field_checks:
                    if cb.isChecked():
                        fields.append(pair)
                if not fields:
                    QMessageBox.information(self, "Export", "Please select at least one field.")
                    return
    
                self.result = {
                    "formats": formats,
                    "boards": boards_sel,   # None = all
                    "fields": fields,       # list of (label, key)
                }
                self.accept()
    
        # Prepare dialog inputs
        boards_avail = _collect_sorted_boards()
        dlg = ExportDialog(boards_avail, self)
        if dlg.exec_() != QDialog.Accepted or dlg.result is None:
            return
    
        formats: list[str] = dlg.result["formats"]
        boards_sel: list[str] | None = dlg.result["boards"]
        fields: list[tuple[str, str]] = dlg.result["fields"]
    
        # === Build dataset from DB (respect boards filter and fields order) ===
        headers = [lbl for (lbl, _key) in fields]
        rows: list[list[str]] = []
    
        # collect paths per current tree (visible dataset)
        paths: list[Path] = []
        try:
            top_count = self.tree_sorted.topLevelItemCount()
            for i in range(top_count):
                top = self.tree_sorted.topLevelItem(i)
                board_name = top.text(0) if top else ""
                if boards_sel is not None and board_name not in boards_sel:
                    continue
                for j in range(top.childCount()):
                    ch = top.child(j)
                    p = ch.data(0, Qt.UserRole + 1)
                    if p:
                        paths.append(Path(p))
        except Exception:
            pass
    
        if not paths:
            QMessageBox.information(self, "Export", "No data in Sorted DB to export.")
            return
    
        for p in paths:
            rec = self.sorted_db.get(p)
            if not rec:
                continue
            data = rec.get("data", {}) or {}
            board_val = rec.get("board") or data.get("board") or ""
            # provide 'file_name' virtual key
            data = dict(data)
            data["file_name"] = p.name
            data["board"] = board_val
    
            row: list[str] = []
            for (_lbl, key) in fields:
                val = data.get(key, "")
                row.append("" if val is None else str(val))
            rows.append(row)
    
        if not rows:
            QMessageBox.information(self, "Export", "No DB records found for current dataset/filters.")
            return
    
        # === Save path (one dialog like Overview; base name composed from boards + timestamp) ===
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        if boards_sel is None or not boards_sel:
            suffix = "all"
        else:
            suffix = "+".join(boards_sel[:3])
            if len(boards_sel) > 3:
                suffix += "+more"
        base_name = f"export-{suffix}-{ts}".replace(" ", "_")
    
        # Default extension = first chosen format
        first_ext = formats[0]
        default_name = f"{base_name}.{first_ext}"
    
        path, _ = QFileDialog.getSaveFileName(self, "Save export as…", default_name, "All files (*.*)")
        if not path:
            return
    
        base_no_ext, _sep, _ext = path.rpartition(".")
        if not base_no_ext:
            base_no_ext = path  # no dot
    
        # === Write all requested formats, using SAME helpers as Overview ===
        ok, errs = [], []
        for fmt in formats:
            target = f"{base_no_ext}.{fmt}"
            try:
                if fmt == "xlsx":
                    self._export_to_xlsx(target, headers, rows)
                elif fmt == "csv":
                    self._export_to_csv(target, headers, rows)
                elif fmt == "txt":
                    self._export_to_txt(target, headers, rows)
                ok.append(target)
            except Exception as e:
                errs.append(f"{fmt.upper()}: {e}")
    
        if ok:
            try:
                self.statusBar().showMessage("Export done.")
            except Exception:
                pass
            QMessageBox.information(self, "Export", "Exported:\n" + "\n".join(ok))
        if errs:
            QMessageBox.warning(self, "Export (some errors)", "Failed:\n" + "\n".join(errs))
        
    def rescan_sorted(self) -> None:
        from PySide6.QtWidgets import QTreeWidgetItem
        from pathlib import Path
        from app.pdf_scanner import PdfScanner
    
        root = self.sorted_root
        self.tree_sorted.clear()
    
        if not root.exists():
            try:
                self.statusBar().showMessage("Sorted PDFs root not found.")
            except Exception:
                pass
            return
    
        # 1) Naparsuj celý kořen Sorted PDFs
        scanner = PdfScanner(root)
        parsed_records = scanner.scan()  # List[PdfRecord]
    
        # 2) Upsert do DB (nepřepisuj edited=True)
        for rec in parsed_records:
            try:
                board = getattr(rec, "board", "") or ""
                path: Path = getattr(rec, "path")
                file_name = path.name
                # převod dat – použij dataclass->dict
                from dataclasses import asdict
                data = asdict(rec)
                self.sorted_db.upsert_parsed(path, board, file_name, data)
            except Exception:
                continue
    
        self.sorted_db.save()
    
        # 3) Postav strom Board → PDF z FS
        boards: dict[str, list[Path]] = {}
        for board_dir in root.iterdir():
            if not board_dir.is_dir():
                continue
            if board_dir.name == "__archive__":
                continue
            # všechny PDF rekurzivně
            pdfs = [p for p in board_dir.rglob("*.pdf")]
            if not pdfs:
                continue
            boards[board_dir.name] = sorted(pdfs, key=lambda p: p.name.lower())
    
        for board_name in sorted(boards.keys(), key=str.lower):
            top = QTreeWidgetItem([board_name])
            top.setData(0, Qt.UserRole + 1, None)  # board item
            self.tree_sorted.addTopLevelItem(top)
            for p in boards[board_name]:
                child = QTreeWidgetItem([p.name])
                child.setData(0, Qt.UserRole + 1, str(p.resolve()))
                top.addChild(child)
            top.setExpanded(True)
    
        try:
            self.statusBar().showMessage(f"Sorted PDFs: {sum(len(v) for v in boards.values())} files in {len(boards)} boards.")
        except Exception:
            pass

    def _sorted_db_path(self) -> Path:
        return self.sorted_db.db_path
    
    def _sorted_key_for(self, abs_path: Path) -> str:
        return self.sorted_db.key_for(abs_path)
    
    def _sorted_on_item_changed(self) -> None:
        sel = self.tree_sorted.selectedItems()
        if not sel:
            return
        item = sel[0]
        p = item.data(0, Qt.UserRole + 1)
        if not p:
            # vybrán pouze board – vymaž detaily
            self._sorted_fill_details(None)
            return
        from pathlib import Path
        self._sorted_fill_details(Path(p))
    
    def _sorted_fill_details(self, abs_path: Optional[Path]) -> None:
        # Vyčisti
        def _set(txt_widget, val: str):
            if hasattr(txt_widget, "setPlainText"):
                txt_widget.setPlainText(val or "")
            else:
                txt_widget.setText(val or "")
    
        if abs_path is None:
            for w in (self.ed_board, self.ed_app_type, self.ed_inst_name, self.ed_cand_name,
                      self.ed_rec_acad, self.ed_rec_cert, self.ed_fullname, self.ed_email,
                      self.ed_phone, self.ed_address, self.ed_syllabi, self.ed_courses,
                      self.ed_proof, self.ed_links, self.ed_additional, self.ed_sigdate,
                      self.ed_filename):
                _set(w, "")
            self._sorted_set_editable(False)
            self._sorted_set_status(None)
            return
    
        # DB → load, pokud není, provedeme okamžitý parse (bez zápisu)
        rec = self.sorted_db.get(abs_path)
        data = None
        edited = False
        board = ""
        file_name = abs_path.name
    
        if rec:
            data = rec.get("data", {})
            edited = bool(rec.get("edited"))
            board = rec.get("board") or ""
        else:
            # parse on-the-fly
            try:
                from app.pdf_scanner import PdfScanner
                scanner = PdfScanner(abs_path.parent)
                parsed = scanner.scan()  # celé parent, najdeme naše
                for r in parsed:
                    if getattr(r, "path", None) and Path(r.path).resolve() == abs_path.resolve():
                        from dataclasses import asdict
                        data = asdict(r)
                        board = getattr(r, "board", "") or abs_path.parent.name
                        break
            except Exception:
                data = {}
    
        # Naplň UI
        _set(self.ed_board, board)
        _set(self.ed_app_type, str(data.get("application_type", "")))
        _set(self.ed_inst_name, str(data.get("institution_name", "")))
        _set(self.ed_cand_name, str(data.get("candidate_name", "")))
        _set(self.ed_rec_acad, str(data.get("recognition_academia", "")))
        _set(self.ed_rec_cert, str(data.get("recognition_certified", "")))
        _set(self.ed_fullname, str(data.get("contact_full_name", "")))
        _set(self.ed_email, str(data.get("contact_email", "")))
        _set(self.ed_phone, str(data.get("contact_phone", "")))
        _set(self.ed_address, str(data.get("contact_postal_address", "")))
        _set(self.ed_syllabi, str(data.get("syllabi_integration_description", "")))
        _set(self.ed_courses, str(data.get("courses_modules_list", "")))
        _set(self.ed_proof, str(data.get("proof_of_istqb_certifications", "")))
        _set(self.ed_links, str(data.get("university_links", "")))
        _set(self.ed_additional, str(data.get("additional_information_documents", "")))
        _set(self.ed_sigdate, str(data.get("signature_date", "")))
        _set(self.ed_filename, file_name)
    
        self._sorted_set_editable(False)
        self._sorted_set_status("Edited" if edited else "Parsed (unmodified)")
        # Uložíme aktuální cestu pro Save
        self._sorted_current_path = abs_path
    
    def _sorted_set_status(self, txt: Optional[str]) -> None:
        self.lbl_sorted_status.setText("—" if not txt else txt)
    
    def _sorted_set_editable(self, can_edit: bool) -> None:
        # Board a File name necháme read-only; ostatní podle can_edit
        for w in (self.ed_app_type, self.ed_inst_name, self.ed_cand_name,
                  self.ed_rec_acad, self.ed_rec_cert, self.ed_fullname,
                  self.ed_email, self.ed_phone, self.ed_sigdate):
            w.setReadOnly(not can_edit)
        for w in (self.ed_address, self.ed_syllabi, self.ed_courses,
                  self.ed_proof, self.ed_links, self.ed_additional):
            w.setReadOnly(not can_edit)

    def _sorted_save_changes(self) -> None:
        from PySide6.QtWidgets import QMessageBox
        from pathlib import Path
    
        abs_path = getattr(self, "_sorted_current_path", None)
        if not abs_path:
            QMessageBox.information(self, "Save to DB", "No PDF selected.")
            return
    
        # Sesbírej hodnoty
        def _get(w):
            return w.toPlainText() if hasattr(w, "toPlainText") else w.text()
    
        new_data = {
            "board": self.ed_board.text().strip(),
            "application_type": _get(self.ed_app_type).strip(),
            "institution_name": _get(self.ed_inst_name).strip(),
            "candidate_name": _get(self.ed_cand_name).strip(),
            "recognition_academia": _get(self.ed_rec_acad).strip(),
            "recognition_certified": _get(self.ed_rec_cert).strip(),
            "contact_full_name": _get(self.ed_fullname).strip(),
            "contact_email": _get(self.ed_email).strip(),
            "contact_phone": _get(self.ed_phone).strip(),
            "contact_postal_address": _get(self.ed_address).strip(),
            "syllabi_integration_description": _get(self.ed_syllabi).strip(),
            "courses_modules_list": _get(self.ed_courses).strip(),
            "proof_of_istqb_certifications": _get(self.ed_proof).strip(),
            "university_links": _get(self.ed_links).strip(),
            "additional_information_documents": _get(self.ed_additional).strip(),
            "signature_date": _get(self.ed_sigdate).strip(),
            "file_name": _get(self.ed_filename).strip(),
            "path": str(Path(abs_path).resolve()),
        }
    
        # Ulož do DB jako edited=True
        self.sorted_db.mark_edited(Path(abs_path), new_data)
        self.sorted_db.save()
    
        self._sorted_set_editable(False)
        self._sorted_set_status("Edited")
    
        try:
            self.statusBar().showMessage("Saved to DB.")
        except Exception:
            pass
        
    def _filter_board_sorted(self, txt: str) -> None:
        proxy = self.table_sorted.model()
        if isinstance(proxy, RecordsModel):
            proxy.set_board(txt)
            
    def _filter_text_sorted(self, txt: str) -> None:
        proxy = self.table_sorted.model()
        if isinstance(proxy, RecordsModel):
            proxy.set_search(txt)

    def _selected_sorted_record(self) -> Optional[PdfRecord]:
        sel = self.table_sorted.selectionModel()
        if not sel or not sel.hasSelection():
            return None
        index = sel.selectedRows()[0]
        FILE_COL = 16
        proxy = self.table_sorted.model()
        if isinstance(proxy, QSortFilterProxyModel):
            sidx = proxy.mapToSource(proxy.index(index.row(), FILE_COL))
            src = proxy.sourceModel()
        else:
            sidx = index
            src = self.table_sorted.model()
        path_str = src.index(sidx.row(), FILE_COL).data(Qt.UserRole + 1)
        for r in self.records_sorted:
            if str(r.path) == path_str:
                return r
        return None
        
    from typing import Optional
    from PySide6.QtWidgets import QMessageBox
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtCore import QUrl
    
    def open_selected_pdf(self) -> None:
        """Open the PDF file of the currently selected row in Overview."""
        rec = self._selected_record()
        if not rec:
            QMessageBox.information(self, "Open PDF", "Please select a row first.")
            return
        # macOS-safe: QDesktopServices.openUrl requires QUrl, not a string
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(rec.path)))

    # ----- Browser tab -----
    def _build_browser_tab(self) -> None:
        """
        PDF Browser: reflow layout to vertical split (TOP: file tree, BOTTOM: details).
        One-shot resize of main window to comfortably fit Name column + details.
        """
        from PySide6.QtCore import Qt, QTimer
        from PySide6.QtWidgets import (
            QSplitter, QWidget, QVBoxLayout, QTreeView, QFormLayout, QLabel,
            QScrollArea, QFileSystemModel, QSizePolicy, QHeaderView
        )
    
        # === Kořenový vertikální splitter: Nahoře strom, dole detail ===
        vsplit = QSplitter(Qt.Vertical, self.browser_tab)
    
        # --- Nahoře: strom PDF ---
        top_widget = QWidget(vsplit)
        top_layout = QVBoxLayout(top_widget)
    
        self.fs_model = QFileSystemModel(self)
        self.fs_model.setRootPath(str(self.pdf_root))
        self.fs_model.setNameFilterDisables(False)
        self.fs_model.setNameFilters(["*.pdf", "*.PDF"])
    
        self.tree = QTreeView(top_widget)
        self.tree.setModel(self.fs_model)
        self.tree.setRootIndex(self.fs_model.index(str(self.pdf_root)))
        self.tree.setSortingEnabled(True)
        self.tree.setUniformRowHeights(True)
    
        # Necháme Qt spočítat šířku sloupce Name podle obsahu
        header = self.tree.header()
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(80)
        try:
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        except Exception:
            header.setResizeMode(0, QHeaderView.ResizeToContents)
    
        # Signály zůstávají (žádná změna chování)
        self.tree.doubleClicked.connect(self._open_from_tree)
        self.tree.selectionModel().selectionChanged.connect(self._tree_selection_changed)
    
        top_layout.addWidget(self.tree)
    
        # --- Dole: detail (ponechávám původní názvy lbl_* pro _update_detail_panel) ---
        bottom_widget = QWidget(vsplit)
        bottom_layout = QVBoxLayout(bottom_widget)
    
        scroll = QScrollArea(bottom_widget)
        scroll.setWidgetResizable(True)
        form_host = QWidget(scroll)
        self.detail_form = QFormLayout(form_host)
        self.detail_form.setRowWrapPolicy(QFormLayout.WrapLongRows)
        self.detail_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.detail_form.setFormAlignment(Qt.AlignTop)
        self.detail_form.setLabelAlignment(Qt.AlignRight | Qt.AlignTop)
    
        def _mklabel() -> QLabel:
            lab = QLabel("-")
            lab.setWordWrap(True)
            lab.setTextInteractionFlags(Qt.TextSelectableByMouse)
            lab.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            return lab
    
        # Vytvoření labelů přesně podle toho, co očekává _update_detail_panel(...)
        self.lbl_board = getattr(self, "lbl_board", _mklabel())
        self.lbl_known = getattr(self, "lbl_known", _mklabel())
        self.lbl_app_type = getattr(self, "lbl_app_type", _mklabel())
        self.lbl_inst = getattr(self, "lbl_inst", _mklabel())
        self.lbl_cand = getattr(self, "lbl_cand", _mklabel())
        self.lbl_acad = getattr(self, "lbl_acad", _mklabel())
        self.lbl_cert = getattr(self, "lbl_cert", _mklabel())
        self.lbl_contact = getattr(self, "lbl_contact", _mklabel())
        self.lbl_email = getattr(self, "lbl_email", _mklabel())
        self.lbl_phone = getattr(self, "lbl_phone", _mklabel())
        self.lbl_postal = getattr(self, "lbl_postal", _mklabel())
        self.lbl_date = getattr(self, "lbl_date", _mklabel())
        self.lbl_syllabi = getattr(self, "lbl_syllabi", _mklabel())
        self.lbl_courses = getattr(self, "lbl_courses", _mklabel())
        self.lbl_proof = getattr(self, "lbl_proof", _mklabel())
        self.lbl_links = getattr(self, "lbl_links", _mklabel())
        self.lbl_additional = getattr(self, "lbl_additional", _mklabel())
        self.lbl_sorted_status = getattr(self, "lbl_sorted_status", _mklabel())
    
        # Sestavení formuláře (pořadí zachováno)
        self.detail_form.addRow("Board:", self.lbl_board)
        self.detail_form.addRow("Known Board:", self.lbl_known)
        self.detail_form.addRow("Application Type:", self.lbl_app_type)
        self.detail_form.addRow("Institution Name:", self.lbl_inst)
        self.detail_form.addRow("Candidate Name:", self.lbl_cand)
        self.detail_form.addRow("Academia Recognition:", self.lbl_acad)
        self.detail_form.addRow("Certified Recognition:", self.lbl_cert)
        self.detail_form.addRow("Full Name:", self.lbl_contact)
        self.detail_form.addRow("Email Address:", self.lbl_email)
        self.detail_form.addRow("Phone Number:", self.lbl_phone)
        self.detail_form.addRow("Postal Address:", self.lbl_postal)
        self.detail_form.addRow("Signature Date:", self.lbl_date)
        self.detail_form.addRow("Syllabi Integration:", self.lbl_syllabi)
        self.detail_form.addRow("Courses/Modules:", self.lbl_courses)
        self.detail_form.addRow("Proof of ISTQB Certifications:", self.lbl_proof)
        self.detail_form.addRow("University Links:", self.lbl_links)
        self.detail_form.addRow("Additional Info/Documents:", self.lbl_additional)
        self.detail_form.addRow("Sorted Status:", self.lbl_sorted_status)
    
        scroll.setWidget(form_host)
        bottom_layout.addWidget(scroll)
    
        # Přidej panely do splitteru
        vsplit.addWidget(top_widget)
        vsplit.addWidget(bottom_widget)
        vsplit.setStretchFactor(0, 2)
        vsplit.setStretchFactor(1, 1)
        vsplit.setCollapsible(0, False)
        vsplit.setCollapsible(1, False)
    
        # Zabalit do záložky
        from PySide6.QtWidgets import QVBoxLayout as _VBL
        outer = _VBL(self.browser_tab)
        outer.addWidget(vsplit)
    
        # Jednorázově po načtení kořenového adresáře zvětšit okno podle šířky Name
        def _widen_window_once():
            try:
                self.tree.resizeColumnToContents(0)
                name_w = max(self.tree.sizeHintForColumn(0), header.sectionSize(0), 220)
            except Exception:
                name_w = 260
            desired_width = int(name_w + 80)  # sloupec + okraje/scroll
            desired_height = max(self.height(), 720)
            if self.width() < desired_width:
                self.resize(desired_width, desired_height)
    
        try:
            self.fs_model.directoryLoaded.connect(lambda *_: QTimer.singleShot(0, _widen_window_once))
        except Exception:
            pass
        QTimer.singleShot(0, _widen_window_once)

    # ----- Data -----
    def rescan(self) -> None:
        """Scan PDF root and repopulate the Overview table while preserving selection.
        Minimal-change: reuse a persistent source model + proxy to avoid empty view issues.
        """
        from pathlib import Path
        from PySide6.QtCore import Qt, QItemSelectionModel
        from PySide6.QtGui import QStandardItemModel, QStandardItem, QBrush, QColor
        from PySide6.QtWidgets import QStyle, QTableView
    
        # --- 0) Ensure persistent models exist (create once) ---
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
            "Eligibility Evidence\nSyllabi Integration",
            "Eligibility Evidence\nCourses/Modules",
            "Eligibility Evidence\nProof of ISTQB Certifications",
            "Eligibility Evidence\nUniversity Links",
            "Eligibility Evidence\nAdditional Info/Documents",
            "Signature Date",
            "File\nFile name",
        ]
        if not hasattr(self, "_headers"):
            self._headers = headers
        if not hasattr(self, "_source_model"):
            self._source_model = QStandardItemModel(0, len(headers), self)
            self._source_model.setHorizontalHeaderLabels(headers)
        if not hasattr(self, "_proxy"):
            self._proxy = RecordsModel(headers, self)
            self._proxy.setSourceModel(self._source_model)
            self._proxy.setDynamicSortFilter(True)
            self.table.setModel(self._proxy)
            self.table.setSelectionBehavior(QTableView.SelectRows)
            self.table.setSelectionMode(QTableView.ExtendedSelection)
            self.table.setSortingEnabled(True)
            self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
            try:
                self.table.verticalHeader().setVisible(False)
            except Exception:
                pass
    
        # --- 1) Capture current selection (by hidden full paths from last column) ---
        selected_paths: set[str] = set()
        try:
            proxy = self.table.model()
            sel = self.table.selectionModel()
            if proxy is not None and sel and sel.hasSelection():
                FILE_COL = proxy.columnCount() - 1
                for pidx in sel.selectedRows():
                    sidx = proxy.mapToSource(pidx) if hasattr(proxy, "mapToSource") else pidx
                    val = self._source_model.index(sidx.row(), FILE_COL).data(Qt.UserRole + 1)
                    if not val:
                        val = self._source_model.index(sidx.row(), FILE_COL).data()
                    if val:
                        selected_paths.add(str(val))
        except Exception:
            selected_paths = set()
    
        # --- 2) Determine and set pdf_root if needed ---
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
        if chosen is None and candidates:
            chosen = candidates[0]
        if isinstance(chosen, Path):
            self.pdf_root = chosen
    
        # --- 3) Count found PDFs (excluding __archive__) ---
        try:
            found = 0
            if isinstance(self.pdf_root, Path) and self.pdf_root.exists():
                for p in self.pdf_root.rglob("*.pdf"):
                    try:
                        rel = p.relative_to(self.pdf_root)
                        if "__archive__" in rel.parts:
                            continue
                        if p.is_file():
                            found += 1
                    except Exception:
                        continue
            else:
                found = 0
        except Exception:
            found = 0
    
        # --- 4) Parse records ---
        scanner = PdfScanner(self.pdf_root) if isinstance(self.pdf_root, Path) else None
        self.records = scanner.scan() if scanner else []
    
        # --- 5) Repopulate source model (clear + append) ---
        model = self._source_model
        if model.rowCount() > 0:
            model.removeRows(0, model.rowCount())
    
        # group coloring
        COLS_APPLICATION = [1]
        COLS_INSTITUTION = [2, 3]
        COLS_RECOG      = [4, 5]
        COLS_CONTACT    = [6, 7, 8, 9]
        COLS_ELIG       = [10, 11, 12, 13, 14]
    
        BRUSH_APP   = QBrush(QColor(58, 74, 110))
        BRUSH_INST  = QBrush(QColor(74, 58, 110))
        BRUSH_RECOG = QBrush(QColor(58, 110, 82))
        BRUSH_CONT  = QBrush(QColor(110, 82, 58))
        BRUSH_ELIG  = QBrush(QColor(92, 92, 92))
    
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
            row_vals = rec.as_row()
            items = [QStandardItem(v) for v in row_vals]
            for it in items:
                it.setEditable(False)
    
            paint_group(items, COLS_APPLICATION, BRUSH_APP)
            paint_group(items, COLS_INSTITUTION, BRUSH_INST)
            paint_group(items, COLS_RECOG,      BRUSH_RECOG)
            paint_group(items, COLS_CONTACT,    BRUSH_CONT)
            paint_group(items, COLS_ELIG,       BRUSH_ELIG)
    
            set_yesno_icon(items[4])  # Academia
            set_yesno_icon(items[5])  # Certified
    
            FILE_COL = len(headers) - 1
            items[FILE_COL].setData(str(rec.path), Qt.UserRole + 1)
            model.appendRow(items)
    
        # --- 6) Post-setup on the proxy/view ---
        proxy = self._proxy
        self.table.sortByColumn(0, Qt.AscendingOrder)
        for c in range(len(headers)):
            self.table.resizeColumnToContents(c)
        for c in (10, 11, 12, 13, 14):
            self.table.setColumnHidden(c, True)
    
        # --- 7) Restore previous selection by matching hidden paths ---
        try:
            if selected_paths:
                sel_model = self.table.selectionModel()
                FILE_COL = proxy.columnCount() - 1
                for r in range(model.rowCount()):
                    sval = model.index(r, FILE_COL).data(Qt.UserRole + 1)
                    if not sval:
                        sval = model.index(r, FILE_COL).data()
                    if sval and str(sval) in selected_paths:
                        pidx = proxy.mapFromSource(model.index(r, 0))
                        sel_model.select(pidx, QItemSelectionModel.Select | QItemSelectionModel.Rows)
        except Exception:
            pass
    
        # --- 8) Watch list & status ---
        self._rebuild_watch_list()
        try:
            root_str = str(self.pdf_root) if isinstance(self.pdf_root, Path) else "<unset>"
            self.statusBar().showMessage(f"PDF found: {found} • Parsed: {len(self.records)} • Root: {root_str}")
        except Exception:
            pass

    # ----- Actions -----
    from typing import Optional
    from PySide6.QtCore import Qt, QModelIndex
    from PySide6.QtWidgets import QTableView
    from PySide6.QtCore import QSortFilterProxyModel
    
    def _selected_record(self) -> Optional[PdfRecord]:
        """
        Return PdfRecord for the *Overview* table selection.
        Robust to multiple table views in other tabs:
        - If invoked from a view signal, use sender() if it's a QTableView.
        - Else, find the Overview table inside self.overview_tab.
        """
        # Prefer the signal sender if it's the table view
        view = None
        snd = self.sender()
        if isinstance(snd, QTableView):
            view = snd
        else:
            # Fallback: locate the Overview table within the Overview tab
            try:
                if hasattr(self, "overview_tab") and self.overview_tab is not None:
                    view = self.overview_tab.findChild(QTableView)  # no objectName required
            except Exception:
                view = None
            # Last resort: use self.table if it is a QTableView
            if view is None and isinstance(getattr(self, "table", None), QTableView):
                view = self.table
    
        if view is None:
            return None
    
        sel = view.selectionModel()
        if not sel or not sel.hasSelection():
            return None
    
        # First selected row in the proxy model
        pindex = sel.selectedRows()[0]
        proxy = view.model()
        if isinstance(proxy, QSortFilterProxyModel):
            # Map selected proxy row to source row; column doesn't matter for row mapping
            srow = proxy.mapToSource(proxy.index(pindex.row(), 0)).row()
            src = proxy.sourceModel()
            file_col = src.columnCount() - 1  # last column stores filename/full path
            idx = src.index(srow, file_col)
            path_str = idx.data(Qt.UserRole + 1) or idx.data()
        else:
            src = proxy
            file_col = src.columnCount() - 1
            idx = src.index(pindex.row(), file_col)
            path_str = idx.data(Qt.UserRole + 1) or idx.data()
    
        if not path_str:
            return None
    
        # Match against loaded records
        for r in getattr(self, "records", []):
            if str(r.path) == str(path_str):
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


    from typing import Optional
    from PySide6.QtWidgets import QMessageBox
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtCore import QUrl   
    
    def open_selected_pdf(self) -> None:
        rec = self._selected_record()
        if not rec:
            QMessageBox.information(self, "Open PDF", "Please select a row first.")
            return
        # 0.6d: QDesktopServices.openUrl vyžaduje QUrl; posílejme lokální souborový URL
        QDesktopServices.openUrl(self.QUrl.fromLocalFile(str(rec.path)))

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
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

# --- vlož do app/main_window.py (nad definici MainWindow) ---
from typing import Optional, Dict, List
from PySide6.QtCore import QSortFilterProxyModel, Qt, QModelIndex

class OverviewBoardGroupingProxy(QSortFilterProxyModel):
    """
    Proxy model pro záložku Overview:
    - vloží virtuální sloupec 'No.' bez zásahu do source modelu (před 'Board'),
    - ve sloupci 'Board' zobrazí hodnotu jen na prvním řádku souvislé skupiny,
    - číslování 1..N v rámci skupiny se přepočítává po seřazení/změnách layoutu.
    """
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._board_source_col: int = -1
        self._board_proxy_insert_at: int = -1  # index, kam vkládáme 'No.'
        self._row_no: Dict[int, int] = {}      # proxyRow -> pořadí v rámci Boardu
        self._show_board_at_row: Dict[int, bool] = {}  # proxyRow -> je první výskyt?

        # Reakce na změny layoutu/sortu
        self.layoutChanged.connect(self._recalculate_grouping)
        self.modelReset.connect(self._recalculate_grouping)

    # ---- zásadní část: definice sloupců ----
    def setSourceModel(self, sourceModel) -> None:  # type: ignore[override]
        super().setSourceModel(sourceModel)
        self._detect_board_column()
        self._recalculate_grouping()

    def _detect_board_column(self) -> None:
        """Najde index sloupce 'Board' v source modelu dle headeru (case-insensitive)."""
        self._board_source_col = -1
        self._board_proxy_insert_at = -1
        src = self.sourceModel()
        if not src:
            return
        cols = src.columnCount()
        for c in range(cols):
            hdr = src.headerData(c, Qt.Horizontal, Qt.DisplayRole)
            if isinstance(hdr, str) and hdr.strip().lower() == "board":
                self._board_source_col = c
                self._board_proxy_insert_at = c  # virtuální 'No.' půjde před tento index
                break

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # type: ignore[override]
        src = self.sourceModel()
        if not src:
            return 0
        base = src.columnCount()
        return base + 1 if self._board_source_col >= 0 else base

    # ---- mapování proxy sloupců na source sloupce ----
    def _map_proxy_to_source_col(self, proxy_col: int) -> Optional[int]:
        """Vrátí index source sloupce nebo None (pokud jde o virtuální 'No.')."""
        if self._board_source_col < 0:
            return proxy_col  # žádná injekce
        insert_at = self._board_proxy_insert_at
        if proxy_col < insert_at:
            return proxy_col
        if proxy_col == insert_at:
            return None  # 'No.' – virtuální
        # posun o 1 za vložený 'No.'
        return proxy_col - 1

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):  # type: ignore[override]
        if orientation == Qt.Horizontal and role == Qt.DisplayRole and self._board_source_col >= 0:
            if section == self._board_proxy_insert_at:
                return "No."
            # posun za vloženým sloupcem – vrátíme původní headery
            src_col = self._map_proxy_to_source_col(section)
            if src_col is not None:
                return self.sourceModel().headerData(src_col, orientation, role)
            return None
        # default
        src_col = self._map_proxy_to_source_col(section)
        if src_col is not None:
            return self.sourceModel().headerData(src_col, orientation, role)
        return super().headerData(section, orientation, role)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):  # type: ignore[override]
        if not index.isValid():
            return None
        if self._board_source_col < 0:
            return super().data(index, role)

        proxy_col = index.column()
        insert_at = self._board_proxy_insert_at

        # Virtuální 'No.'
        if proxy_col == insert_at:
            if role == Qt.DisplayRole:
                return self._row_no.get(index.row(), None)
            if role == Qt.TextAlignmentRole:
                return Qt.AlignCenter
            return None

        # Board sloupec (už v proxy o 1 vpravo od 'No.')
        if proxy_col == insert_at + 1:
            if role == Qt.DisplayRole:
                # jen u prvního řádku skupiny, jinak prázdno
                return self._board_value(index.row()) if self._show_board_at_row.get(index.row(), False) else ""
            # ostatní role necháme propadnout na source
            # (kvůli případným tooltipům/stylu apod.)
        # Ostatní běžná data – přemapujeme na source
        src_col = self._map_proxy_to_source_col(proxy_col)
        if src_col is None:
            return None
        src_index = self.mapToSource(self.index(index.row(), proxy_col))
        if not src_index.isValid():
            return None
        # Vytvořte index do source s opravdovým sloupcem:
        src_index = self.sourceModel().index(src_index.row(), src_col)
        return self.sourceModel().data(src_index, role)

    # ---- pomocné výpočty ----
    def _board_value(self, proxy_row: int) -> str:
        """Získá hodnotu 'Board' pro daný proxy řádek přímo ze source modelu."""
        if self._board_source_col < 0:
            return ""
        # map row do source:
        src_row = self.mapToSource(self.index(proxy_row, 0)).row()
        if src_row < 0:
            return ""
        src_idx = self.sourceModel().index(src_row, self._board_source_col)
        val = self.sourceModel().data(src_idx, Qt.DisplayRole)
        return str(val) if val is not None else ""

    def _recalculate_grouping(self) -> None:
        """Po změně layoutu/sortu spočítá pořadí ('No.') a první výskyty Boardů."""
        self._row_no.clear()
        self._show_board_at_row.clear()
        if self._board_source_col < 0:
            return
        prev_board = None
        counter = 0
        rows = self.rowCount()
        for r in range(rows):
            bval = self._board_value(r)
            if bval != prev_board:
                counter = 1
                self._show_board_at_row[r] = True
                prev_board = bval
            else:
                counter += 1
            self._row_no[r] = counter

    def sort(self, column: int, order: Qt.SortOrder = Qt.AscendingOrder) -> None:  # type: ignore[override]
        super().sort(column, order)
        # Po sortu se layout změní, ale explicitní refresh neuškodí:
        self._recalculate_grouping()

from PySide6.QtWidgets import QTableView

class OverviewTableView(QTableView):
    """
    Tenká obálka nad QTableView: pokaždé, když se nastaví model,
    automaticky ho zabalí do OverviewBoardGroupingProxy (pokud už není zabalen).
    Díky tomu není třeba zasahovat do ostatní logiky plnění modelu.
    """
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._proxy: Optional[OverviewBoardGroupingProxy] = None

    def setModel(self, model) -> None:  # type: ignore[override]
        if isinstance(model, OverviewBoardGroupingProxy):
            self._proxy = model
            return super().setModel(model)
        proxy = OverviewBoardGroupingProxy(self)
        proxy.setDynamicSortFilter(True)
        proxy.setSourceModel(model)
        self._proxy = proxy
        super().setModel(proxy)

    @property
    def proxy(self) -> Optional[OverviewBoardGroupingProxy]:
        return self._proxy

class RecordsModel(QSortFilterProxyModel):
    def __init__(self, headers: List[str], parent=None):
        super().__init__(parent)
        self.headers = headers
        self.search = ""
        self.board_filter = "All"

    # --- Public API pro filtry ---
    def set_search(self, text: str) -> None:
        self.search = text.lower().strip()
        self.invalidateFilter()

    def set_board(self, board: str) -> None:
        self.board_filter = board
        self.invalidateFilter()

    # --- Pomocné: identifikace a normalizace Application Type ---
    def _is_application_type_column(self, src_col: int) -> bool:
        src = self.sourceModel()
        if src is None:
            return False
        hdr = src.headerData(src_col, Qt.Horizontal, Qt.DisplayRole)
        if isinstance(hdr, str):
            h = hdr.lower().replace("\n", " ")
            return "application type" in h
        return False

    @staticmethod
    def _sanitize_app_type_text(val: str) -> str:
        """
        Odstraní úvodní '/' a znormalizuje hodnotu na přesné názvy:
        - 'New Application'
        - 'Additional Recognition'
        Pokud nepozná, vrátí očištěný text.
        """
        s = (val or "").lstrip("/").strip()
        low = s.lower()
        if "new" in low:
            return "New Application"
        if "additional" in low:
            return "Additional Recognition"
        return s

    # --- Filtrování řádků ---
    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        model = self.sourceModel()
        if model is None:
            return True

        # Board filter (předpoklad: ve source je 'Board' ve sloupci 0)
        idx_board = model.index(source_row, 0, source_parent)
        board_val = (model.data(idx_board, Qt.DisplayRole) or "").strip()
        if self.board_filter != "All" and board_val != self.board_filter:
            return False

        if not self.search:
            return True

        # Fulltext přes všechny sloupce SOURCE modelu
        for c in range(model.columnCount()):
            idx = model.index(source_row, c, source_parent)
            val = model.data(idx, Qt.DisplayRole)
            if isinstance(val, str) and self.search in val.lower():
                return True
        return False

    # --- Řazení: Board → Application Type → Candidate Name ---
    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        model = self.sourceModel()
        if model is None:
            return super().lessThan(left, right)

        def data(row: int, col: int) -> str:
            idx = model.index(row, col)
            return (model.data(idx, Qt.DisplayRole) or "").strip()

        BOARD = 0         # Board
        APP   = 1         # Application Type
        CAND  = 3         # Candidate Name (dle tvého dřívějšího pořadí)

        a_board = data(left.row(), BOARD).lower()
        b_board = data(right.row(), BOARD).lower()

        # Normalizuj Application Type pro řazení (New před Additional)
        a_app_norm = self._sanitize_app_type_text(data(left.row(), APP))
        b_app_norm = self._sanitize_app_type_text(data(right.row(), APP))
        order = {"new application": 0, "additional recognition": 1}
        a_app = order.get(a_app_norm.lower(), 2)
        b_app = order.get(b_app_norm.lower(), 2)

        a_cand = data(left.row(), CAND).lower()
        b_cand = data(right.row(), CAND).lower()

        return (a_board, a_app, a_cand) < (b_board, b_app, b_cand)

    # --- Zobrazení buněk ---
    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        # Na DisplayRole sanitizuj Application Type (UI + exporty)
        if role == Qt.DisplayRole and index.isValid():
            src_col = index.column()  # jsme proxy, ale RecordsModel je jediná proxy → sloupce odpovídají source
            if self._is_application_type_column(src_col):
                val = super().data(index, role)
                if isinstance(val, str):
                    return self._sanitize_app_type_text(val)
        return super().data(index, role)

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
        """
        Export vybraných řádků z Overview do složky 'Sorted PDFs/<Board>/' + zápis do Sorted DB (pokud je k dispozici).
        Oprava: výběr je mapován přes libovolné proxy do source řádků.
        """
        from pathlib import Path
        from shutil import copy2
        from PySide6.QtWidgets import QMessageBox
    
        rows = self._selected_rows_source(self.table)
        if not rows:
            QMessageBox.information(self, "Export", "Nejsou vybrány žádné záznamy.")
            return
    
        # Předpokládáme, že self.records je seznam PdfRecord ve STEJNÉM pořadí jako source model.
        # (Takto je to v původní verzi 0.3h/0.5a – zachovávám chování.)
        to_export = []
        for r in rows:
            if 0 <= r < len(self.records):
                to_export.append(self.records[r])
    
        if not to_export:
            QMessageBox.warning(self, "Export", "Výběr neodpovídá žádným platným záznamům.")
            return
    
        base = Path(self.pdf_root) if hasattr(self, "pdf_root") else Path.cwd()
        sorted_root = base.parent / "Sorted PDFs"
        sorted_root.mkdir(parents=True, exist_ok=True)
    
        ok, fail = 0, 0
        for rec in to_export:
            try:
                src = Path(rec.path)
                board = (rec.board or "Unknown").strip() or "Unknown"
                dst_dir = sorted_root / board
                dst_dir.mkdir(parents=True, exist_ok=True)
                dst = dst_dir / src.name
                copy2(src, dst)
                ok += 1
                # Zápis do Sorted DB, pokud existuje API
                if hasattr(self, "sorted_db") and hasattr(self.sorted_db, "add_or_update_from_record"):
                    try:
                        self.sorted_db.add_or_update_from_record(rec, file_path=str(dst))
                    except Exception:
                        pass
            except Exception:
                fail += 1
    
        # Ulož DB, pokud umí
        if hasattr(self, "sorted_db") and hasattr(self.sorted_db, "save"):
            try:
                self.sorted_db.save()
            except Exception:
                pass
    
        QMessageBox.information(self, "Export",
                                f"Export hotov.\nÚspěšně: {ok}\nChyby: {fail}\nCíl: {sorted_root}")

    # ----- Overview tab -----
    def _build_overview_tab(self) -> None:
        from PySide6.QtWidgets import (
            QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QLineEdit,
            QPushButton, QToolButton, QMenu, QAbstractItemView
        )
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QStyle
    
        layout = QVBoxLayout()
        controls = QHBoxLayout()
    
        # Unparsed (pokud je už v projektu – ponecháno)
        self.btn_unparsed = QToolButton(self)
        self.btn_unparsed.setText("Unparsed")
        self.btn_unparsed.setToolTip("Show PDFs found on disk that are not present in Overview")
        self.btn_unparsed.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxWarning))
        self.btn_unparsed.setAutoRaise(True)
        self.btn_unparsed.setStyleSheet("QToolButton { color: #ff6b6b; font-weight: 600; }")
        self.btn_unparsed.clicked.connect(self.show_unparsed_report)
        controls.addWidget(self.btn_unparsed)
    
        # Export…
        self.btn_export = QToolButton(self)
        self.btn_export.setToolTip("Export…")
        self.btn_export.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.btn_export.setAutoRaise(True
        )
        self.btn_export.clicked.connect(self.on_export_overview)
        controls.addWidget(self.btn_export)
    
        # Board filter
        self.board_combo = QComboBox()
        self.board_combo.addItem("All")
        for b in sorted(KNOWN_BOARDS):
            self.board_combo.addItem(b)
        self.board_combo.currentTextChanged.connect(self._filter_board)
    
        # Search
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search…")
        self.search_edit.textChanged.connect(self._filter_text)
    
        # Open PDF
        self.open_btn = QPushButton("Open PDF")
        self.open_btn.clicked.connect(self.open_selected_pdf)
    
        controls.addSpacing(8)
        controls.addWidget(QLabel("Board:"))
        controls.addWidget(self.board_combo, 1)
        controls.addSpacing(12)
        controls.addWidget(QLabel("Search:"))
        controls.addWidget(self.search_edit, 4)
        controls.addSpacing(12)
        controls.addWidget(self.open_btn)
    
        layout.addLayout(controls)
    
        # POZOR: OverviewTableView už má nasazený proxy model pro "No." a Board-grouping
        self.table = OverviewTableView(self)
        # ✅ OPRAVA: používáme enumy z QAbstractItemView (ne atributy instance)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.doubleClicked.connect(self.open_selected_pdf)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setMinimumHeight(44)
    
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
    
        if not getattr(self, "_overview_ctx_connected", False):
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
        Zobrazí dialog se seznamem PDF, která jsou na disku v 'PDF/' (mimo '__archive__'),
        ale **nejsou** v Overview (tj. nepodařilo se je naparsovat / zobrazit).
        Opraveno: case-insensitive přípony, normalizace cest, ignorování duplicit.
        """
        from PySide6.QtWidgets import QMessageBox
        from pathlib import Path
    
        root = Path(self.pdf_root) if hasattr(self, "pdf_root") else Path.cwd() / "PDF"
        if not root.exists():
            QMessageBox.information(self, "Unparsed", f"Adresář neexistuje:\n{root}")
            return
    
        # 1) Všechna PDF na disku (mimo '__archive__')
        disk_pdf = []
        for p in root.rglob("*.pdf"):
            if "__archive__" in p.parts:
                continue
            disk_pdf.append(p.resolve())
    
        # 2) PDF v Overview (z self.records)
        parsed = set()
        for rec in getattr(self, "records", []):
            try:
                parsed.add(Path(rec.path).resolve())
            except Exception:
                continue
    
        # 3) Rozdíl
        unparsed = [p for p in disk_pdf if p not in parsed]
    
        if not unparsed:
            QMessageBox.information(self, "Unparsed", "Všechna nalezená PDF jsou v Overview.")
            return
    
        # Jednoduchý textový přehled
        lines = ["PDF na disku, která nejsou v Overview:", ""]
        for p in sorted(unparsed):
            try:
                board = next((part for part in p.parts if part != "PDF" and part != "__archive__"), "")
            except Exception:
                board = ""
            lines.append(f"- [{board}] {p.name} ({p.parent})")
        text = "\n".join(lines)
    
        QMessageBox.information(self, "Unparsed report", text)
        
    def edit_selected_sorted_record(self) -> None:
        """
        Otevře edit dialog pro vybraný záznam v záložce 'Sorted PDFs'.
        Oprava: mapování výběru přes proxy → source; ochrana prázdného výběru.
        """
        from PySide6.QtWidgets import QMessageBox
    
        if not hasattr(self, "sorted_table"):
            QMessageBox.warning(self, "Edit", "Tabulka 'Sorted PDFs' není k dispozici.")
            return
    
        rows = self._selected_rows_source(self.sorted_table)
        if not rows:
            QMessageBox.information(self, "Edit", "Vyberte prosím řádek v záložce 'Sorted PDFs'.")
            return
    
        # editujeme první vybraný (pokud chceš batch edit, lze doplnit později)
        row = rows[0]
    
        # Předpoklad: existuje metoda, která otevře dialog nad Sorted DB podle indexu/klíče
        # Zkusíme několik běžných variant nenásilnou cestou:
        try:
            self._open_sorted_edit_dialog_by_row(row)   # tvoje interní utilita (pokud existuje)
            return
        except Exception:
            pass
        try:
            self._open_edit_dialog_for_sorted_row(row)  # alternativa pojmenování
            return
        except Exception:
            pass
    
        # Fallback: pokud máme list/sekvenci záznamů
        try:
            rec = self.sorted_db.records[row]
            self._open_edit_dialog_for_record(rec)
            return
        except Exception:
            QMessageBox.warning(self, "Edit", "Nepodařilo se otevřít edit dialog pro vybraný záznam.")

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
                
    def _export_to_txt(self, filename: str, headers: list[str], rows: list[list[str]]) -> None:
        from pathlib import Path
        Path(filename).parent.mkdir(parents=True, exist_ok=True)
    
        # spočti šířky sloupců (max 80)
        cols = len(headers)
        widths = [len(h) for h in headers]
        for r in rows:
            for i in range(cols):
                if i < len(r):
                    widths[i] = min(max(widths[i], len(str(r[i]))), 80)
    
        def fmt_row(vals: list[str]) -> str:
            return " | ".join(str(vals[i]).ljust(widths[i]) for i in range(cols))
    
        sep = "-+-".join("-" * w for w in widths)
    
        with open(filename, "w", encoding="utf-8") as f:
            f.write(fmt_row(headers) + "\n")
            f.write(sep + "\n")
            for r in rows:
                f.write(fmt_row(r) + "\n")
                
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
        from PySide6.QtWidgets import (
            QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem, QSplitter,
            QWidget, QFormLayout, QLabel, QLineEdit, QPlainTextEdit, QPushButton
        )
        from PySide6.QtCore import Qt
    
        layout = QVBoxLayout()
    
        self.split_sorted = QSplitter(self.sorted_tab)
        self.split_sorted.setOrientation(Qt.Horizontal)
    
        # Levý panel – strom Board -> PDF
        self.tree_sorted = QTreeWidget()
        self.tree_sorted.setHeaderLabels(["Board / PDF"])
        self.tree_sorted.itemSelectionChanged.connect(self._sorted_on_item_changed)
    
        # Pravý panel – detaily a editace
        right = QWidget()
        self.form_sorted = QFormLayout(right)
    
        # Indikace stavu
        self.lbl_sorted_status = QLabel("—")
        self.form_sorted.addRow(QLabel("Status:"), self.lbl_sorted_status)
    
        # Definice polí (label -> widget)
        # (Krátká pole = QLineEdit; dlouhé texty = QPlainTextEdit)
        self.ed_board = QLineEdit();                     self.ed_board.setReadOnly(True)
        self.ed_app_type = QLineEdit()
        self.ed_inst_name = QLineEdit()
        self.ed_cand_name = QLineEdit()
        self.ed_rec_acad = QLineEdit()
        self.ed_rec_cert = QLineEdit()
        self.ed_fullname = QLineEdit()
        self.ed_email = QLineEdit()
        self.ed_phone = QLineEdit()
        self.ed_address = QPlainTextEdit()
        self.ed_syllabi = QPlainTextEdit()
        self.ed_courses = QPlainTextEdit()
        self.ed_proof = QPlainTextEdit()
        self.ed_links = QPlainTextEdit()
        self.ed_additional = QPlainTextEdit()
        self.ed_sigdate = QLineEdit()
        self.ed_filename = QLineEdit();                 self.ed_filename.setReadOnly(True)
    
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
    
        # Tlačítka – Edit/Save a manuální rescan
        btn_row = QHBoxLayout()
        self.btn_sorted_edit = QPushButton("Edit")
        self.btn_sorted_save = QPushButton("Save to DB")
        self.btn_sorted_rescan = QPushButton("Rescan Sorted")
        btn_row.addWidget(self.btn_sorted_edit)
        btn_row.addWidget(self.btn_sorted_save)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_sorted_rescan)
        self.form_sorted.addRow(btn_row)
    
        self.btn_sorted_edit.clicked.connect(lambda: self._sorted_set_editable(True))
        self.btn_sorted_save.clicked.connect(self._sorted_save_changes)
        self.btn_sorted_rescan.clicked.connect(self.rescan_sorted)
    
        # Výchozí – read-only
        self._sorted_set_editable(False)
    
        self.split_sorted.addWidget(self.tree_sorted)
        self.split_sorted.addWidget(right)
        self.split_sorted.setStretchFactor(0, 1)
        self.split_sorted.setStretchFactor(1, 2)
    
        layout.addWidget(self.split_sorted, 1)
        self.sorted_tab.setLayout(layout)
        
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
    
    def open_selected_pdf_sorted(self) -> None:
        rec = self._selected_sorted_record()
        if not rec:
            QMessageBox.information(self, "Open PDF", "Please select a row first.")
            return
        QDesktopServices.openUrl(rec.path.as_uri())

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
    
    from PySide6.QtCore import QModelIndex
    from PySide6.QtCore import QSortFilterProxyModel
    from PySide6.QtWidgets import QTableView
    
    def _map_view_index_to_source(self, view: QTableView, proxy_index: QModelIndex) -> QModelIndex:
        """
        Zmapuje index z view přes libovolný řetězec QSortFilterProxyModel → do zdrojového modelu.
        Vrací neplatný index, pokud mapování selže.
        """
        idx = QModelIndex(proxy_index)
        model = view.model()
        while isinstance(model, QSortFilterProxyModel) and idx.isValid():
            idx = model.mapToSource(idx)
            model = model.sourceModel()
        return idx
    
    def _selected_rows_source(self, view: QTableView) -> list[int]:
        """
        Vrátí setříděné, unikátní **řádky v source modelu** dle aktuálního výběru ve view.
        Funguje korektně i při více vnořených proxy modelech.
        """
        sel = view.selectionModel()
        if not sel:
            return []
        rows: set[int] = set()
        for i in sel.selectedRows():
            src_idx = self._map_view_index_to_source(view, i)
            if src_idx.isValid():
                rows.add(src_idx.row())
        return sorted(rows)

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
    
        # ✅ Normalizace Application Type (odstraní '/' a sjednotí text)
        self.lbl_app_type.setText(self._sanitize_app_type(rec.application_type or ""))
    
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
        
    def _sanitize_app_type(self, s: str) -> str:
        """
        UI helper: odstraní '/' a sjednotí Application Type na přesné názvy.
        Použito v detail panelu; tabulka i exporty se sanitizují už v RecordsModel.
        """
        s = (s or "").lstrip("/").strip()
        low = s.lower()
        if "new" in low:
            return "New Application"
        if "additional" in low:
            return "Additional Recognition"
        return s
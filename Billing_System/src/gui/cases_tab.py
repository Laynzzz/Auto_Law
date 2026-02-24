"""Tab 1 — Cases table with inline editing, sorting, and filtering."""

from datetime import date, datetime

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QMenu,
    QMessageBox,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.dataset import COLUMNS, VALID_CASE_STATUSES, load_dataset
from src.services import case_service
from src.gui.dialogs import EditChargeDialog, show_error

# ── Column metadata ──────────────────────────────────────────────
HEADERS = {
    "appearance_date":  "Appearance Date",
    "invoice_number":   "Invoice #",
    "index_number":     "Index #",
    "case_caption":     "Case Caption",
    "court":            "Court",
    "outcome":          "Outcome",
    "case_status":      "Status",
    "charge_amount":    "Charge Amount",
    "invoice_sent_date": "Invoice Sent",
    "paid_status":      "Paid Status",
    "payment_date":     "Payment Date",
    "notes":            "Notes",
}

EDITABLE_COLS = {"outcome", "case_status", "notes"}
READ_ONLY_COLS = set(COLUMNS) - EDITABLE_COLS

SORTED_CASE_STATUSES = sorted(VALID_CASE_STATUSES)

# Column index helpers
COL_INDEX = {col: i for i, col in enumerate(COLUMNS)}
CHARGE_COL = COL_INDEX["charge_amount"]
STATUS_COL = COL_INDEX["case_status"]


# ── StatusDelegate (QComboBox for case_status) ───────────────────

class StatusDelegate(QStyledItemDelegate):
    """Drop-down editor for the case_status column."""

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.addItems(SORTED_CASE_STATUSES)
        return combo

    def setEditorData(self, editor, index):
        value = index.data(Qt.DisplayRole) or ""
        idx = editor.findText(value)
        if idx >= 0:
            editor.setCurrentIndex(idx)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)


# ── CasesTab ─────────────────────────────────────────────────────

class CasesTab(QWidget):
    """Main table displaying all cases for a firm."""

    caseSelected = Signal(object)  # dict | None

    def __init__(self, parent=None):
        super().__init__(parent)
        self._firm: str | None = None
        self._config: dict | None = None
        self._all_rows: list[dict] = []
        self._shown_rows: list[dict] = []
        self._filters: dict = {}
        self._updating = False  # guard against cellChanged during programmatic updates

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._table = QTableWidget()
        self._table.setColumnCount(len(COLUMNS))
        self._table.setHorizontalHeaderLabels(
            [HEADERS[c] for c in COLUMNS]
        )
        self._table.setSortingEnabled(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        self._table.verticalHeader().setVisible(False)

        # Delegate for case_status column
        self._table.setItemDelegateForColumn(
            STATUS_COL, StatusDelegate(self._table)
        )

        # Signals
        self._table.cellChanged.connect(self._on_cell_changed)
        self._table.itemSelectionChanged.connect(self._on_selection_changed)
        self._table.customContextMenuRequested.connect(self._on_context_menu)
        self._table.cellDoubleClicked.connect(self._on_double_click)

        layout.addWidget(self._table)

    # ── public API ────────────────────────────────────────────────

    def load_firm(self, firm: str, config: dict | None = None):
        """Load (or reload) a firm's dataset from disk."""
        self._firm = firm
        self._config = config
        self._all_rows = load_dataset(firm)
        self._apply_and_render()

    def apply_filters(self, filters: dict):
        """Apply in-memory filters and re-render.

        Supported keys: date_from, date_to, case_status, paid_status.
        """
        self._filters = filters
        self._apply_and_render()

    def selected_case(self) -> dict | None:
        row = self._table.currentRow()
        if row < 0 or row >= len(self._shown_rows):
            return None
        return self._shown_rows[row]

    def refresh(self):
        """Re-read from disk and re-apply current filters."""
        if self._firm:
            self._all_rows = load_dataset(self._firm)
            self._apply_and_render()

    # ── filtering ─────────────────────────────────────────────────

    def _apply_and_render(self):
        rows = self._all_rows
        f = self._filters

        date_from = f.get("date_from")
        date_to = f.get("date_to")
        cs = f.get("case_status")
        ps = f.get("paid_status")

        if date_from or date_to or cs or ps:
            filtered = []
            for r in rows:
                ad = self._parse_date(r.get("appearance_date"))
                if date_from and ad and ad < date_from:
                    continue
                if date_to and ad and ad > date_to:
                    continue
                if cs and (r.get("case_status") or "") != cs:
                    continue
                if ps and (r.get("paid_status") or "") != ps:
                    continue
                filtered.append(r)
            rows = filtered

        self._shown_rows = rows
        self._render()

    def _render(self):
        self._updating = True
        self._table.setSortingEnabled(False)
        self._table.setRowCount(len(self._shown_rows))

        for row_idx, row_data in enumerate(self._shown_rows):
            for col_idx, col_name in enumerate(COLUMNS):
                raw = row_data.get(col_name)
                item = QTableWidgetItem()

                if col_name == "charge_amount":
                    val = self._to_float(raw)
                    item.setText(f"${val:,.2f}" if val is not None else "")
                    item.setData(Qt.UserRole, val if val is not None else 0.0)
                elif col_name in ("appearance_date", "invoice_sent_date",
                                  "payment_date"):
                    d = self._parse_date(raw)
                    item.setText(d.isoformat() if d else "")
                else:
                    item.setText(str(raw) if raw is not None and str(raw) != "nan" else "")

                # Editability
                if col_name in EDITABLE_COLS:
                    item.setFlags(
                        item.flags() | Qt.ItemIsEditable
                    )
                else:
                    item.setFlags(
                        item.flags() & ~Qt.ItemIsEditable
                    )

                self._table.setItem(row_idx, col_idx, item)

        self._table.setSortingEnabled(True)
        self._updating = False

    # ── inline edit handling ──────────────────────────────────────

    def _on_cell_changed(self, row: int, col: int):
        if self._updating or not self._firm:
            return
        col_name = COLUMNS[col]
        if col_name not in EDITABLE_COLS:
            return

        case = self._shown_rows[row] if row < len(self._shown_rows) else None
        if not case:
            return

        new_value = (self._table.item(row, col).text() or "").strip()

        result = case_service.edit_case_field(
            firm=self._firm,
            index_number=str(case["index_number"]),
            appearance_date=self._date_str(case["appearance_date"]),
            field_name=col_name,
            new_value=new_value,
            config=self._config,
        )

        if not result.success:
            # Revert cell to old value
            self._updating = True
            old = case.get(col_name)
            self._table.item(row, col).setText(
                str(old) if old is not None and str(old) != "nan" else ""
            )
            self._updating = False
            QMessageBox.warning(self, "Edit failed", result.message)
        else:
            # Update local cache
            case[col_name] = new_value

    # ── context menu / double-click for charge_amount ─────────────

    def _on_context_menu(self, pos):
        item = self._table.itemAt(pos)
        if not item:
            return
        col = item.column()
        if COLUMNS[col] != "charge_amount":
            return
        menu = QMenu(self)
        action = menu.addAction("Edit Charge Amount...")
        action.triggered.connect(lambda: self._edit_charge(item.row()))
        menu.exec(self._table.viewport().mapToGlobal(pos))

    def _on_double_click(self, row: int, col: int):
        if COLUMNS[col] == "charge_amount":
            self._edit_charge(row)

    def _edit_charge(self, row: int):
        if row < 0 or row >= len(self._shown_rows):
            return
        case = self._shown_rows[row]
        current = self._to_float(case.get("charge_amount")) or 0.0
        has_invoice = bool(case.get("invoice_sent_date")
                          and str(case.get("invoice_sent_date")) not in ("", "nan", "NaT"))

        dlg = EditChargeDialog(current, invoice_sent=has_invoice, parent=self)
        if dlg.exec() != EditChargeDialog.Accepted:
            return

        result = case_service.edit_case_field(
            firm=self._firm,
            index_number=str(case["index_number"]),
            appearance_date=self._date_str(case["appearance_date"]),
            field_name="charge_amount",
            new_value=str(dlg.new_amount()),
            reason=dlg.reason() or None,
            config=self._config,
        )

        if result.success:
            self.refresh()
        else:
            show_error(self, "Edit Charge Failed", result.message)

    # ── selection ─────────────────────────────────────────────────

    def _on_selection_changed(self):
        self.caseSelected.emit(self.selected_case())

    # ── helpers ───────────────────────────────────────────────────

    @staticmethod
    def _parse_date(val) -> date | None:
        if val is None:
            return None
        if isinstance(val, date):
            return val
        if isinstance(val, datetime):
            return val.date()
        s = str(val).strip()
        if not s or s in ("nan", "NaT"):
            return None
        for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        return None

    @staticmethod
    def _to_float(val) -> float | None:
        if val is None:
            return None
        try:
            return float(val)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _date_str(val) -> str:
        """Convert a value to YYYY-MM-DD string for service calls."""
        if isinstance(val, (date, datetime)):
            d = val if isinstance(val, date) else val.date()
            return d.isoformat()
        return str(val).strip()

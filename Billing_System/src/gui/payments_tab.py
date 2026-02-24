"""Tab 4 — Payments table for marking invoices paid/unpaid/partial."""

from datetime import date, datetime

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.dataset import load_dataset
from src.services import payment_service
from src.gui.dialogs import MarkPaymentDialog, show_error

# ── Column metadata ──────────────────────────────────────────────

PAY_COLUMNS = [
    "invoice_number",
    "case_caption",
    "appearance_date",
    "charge_amount",
    "paid_status",
    "payment_date",
    "notes",
]

PAY_HEADERS = {
    "invoice_number":   "Invoice #",
    "case_caption":     "Case Caption",
    "appearance_date":  "Appearance Date",
    "charge_amount":    "Charge Amount",
    "paid_status":      "Paid Status",
    "payment_date":     "Payment Date",
    "notes":            "Notes",
}


class PaymentsTab(QWidget):
    """Table of cases with invoice numbers — mark payment via dialog."""

    paymentUpdated = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._firm: str | None = None
        self._config: dict | None = None
        self._rows: list[dict] = []

        layout = QVBoxLayout(self)

        # ── Filter bar ───────────────────────────────────────────
        bar = QHBoxLayout()

        bar.addWidget(QLabel("Paid Status:"))
        self._filter_combo = QComboBox()
        self._filter_combo.addItems(["", "Paid", "Unpaid", "Partial"])
        self._filter_combo.setCurrentIndex(2)  # Default: "Unpaid"
        self._filter_combo.currentTextChanged.connect(self._apply_filter)
        bar.addWidget(self._filter_combo)

        bar.addStretch()

        self._btn_refresh = QPushButton("Refresh")
        self._btn_refresh.clicked.connect(self.refresh)
        bar.addWidget(self._btn_refresh)

        self._btn_mark = QPushButton("Mark Payment...")
        self._btn_mark.clicked.connect(self._mark_selected)
        bar.addWidget(self._btn_mark)

        layout.addLayout(bar)

        # ── Table ────────────────────────────────────────────────
        self._table = QTableWidget()
        self._table.setColumnCount(len(PAY_COLUMNS))
        self._table.setHorizontalHeaderLabels(
            [PAY_HEADERS[c] for c in PAY_COLUMNS]
        )
        self._table.setSortingEnabled(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        self._table.verticalHeader().setVisible(False)

        self._table.cellDoubleClicked.connect(self._on_double_click)
        self._table.customContextMenuRequested.connect(self._on_context_menu)

        layout.addWidget(self._table)

    # ── public API ────────────────────────────────────────────────

    def set_firm(self, firm: str, config: dict | None = None):
        self._firm = firm
        self._config = config
        self._load_data()

    def refresh(self):
        if self._firm:
            self._load_data()

    # ── data loading ──────────────────────────────────────────────

    def _load_data(self):
        """Load dataset, keep only rows with invoice numbers."""
        all_rows = load_dataset(self._firm)
        self._rows = [
            r for r in all_rows
            if r.get("invoice_number")
            and str(r["invoice_number"]).strip()
            and str(r["invoice_number"]).strip() not in ("nan", "None")
        ]
        self._apply_filter()

    def _apply_filter(self, _text: str | None = None):
        status_filter = self._filter_combo.currentText()
        if status_filter:
            shown = [
                r for r in self._rows
                if (r.get("paid_status") or "") == status_filter
            ]
        else:
            shown = list(self._rows)
        self._render(shown)

    def _render(self, rows: list[dict]):
        self._shown_rows = rows
        self._table.setSortingEnabled(False)
        self._table.setRowCount(len(rows))

        for row_idx, row_data in enumerate(rows):
            for col_idx, col_name in enumerate(PAY_COLUMNS):
                raw = row_data.get(col_name)
                item = QTableWidgetItem()

                if col_name == "charge_amount":
                    val = self._to_float(raw)
                    item.setText(f"${val:,.2f}" if val is not None else "")
                    item.setData(Qt.UserRole, val if val is not None else 0.0)
                elif col_name in ("appearance_date", "payment_date"):
                    d = self._parse_date(raw)
                    item.setText(d.isoformat() if d else "")
                else:
                    item.setText(
                        str(raw) if raw is not None and str(raw) not in ("nan", "None") else ""
                    )

                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self._table.setItem(row_idx, col_idx, item)

        self._table.setSortingEnabled(True)

    # ── mark payment ──────────────────────────────────────────────

    def _mark_selected(self):
        row = self._table.currentRow()
        if row < 0:
            QMessageBox.information(self, "No selection", "Select a row first.")
            return
        self._mark_payment(row)

    def _on_double_click(self, row: int, _col: int):
        self._mark_payment(row)

    def _on_context_menu(self, pos):
        item = self._table.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        action = menu.addAction("Mark Payment...")
        action.triggered.connect(lambda: self._mark_payment(item.row()))
        menu.exec(self._table.viewport().mapToGlobal(pos))

    def _mark_payment(self, row: int):
        if row < 0 or row >= len(self._shown_rows):
            return
        case = self._shown_rows[row]
        inv = str(case.get("invoice_number", "")).strip()
        caption = str(case.get("case_caption", ""))
        current_status = str(case.get("paid_status", "")).strip()

        dlg = MarkPaymentDialog(inv, caption, current_status, parent=self)
        if dlg.exec() != MarkPaymentDialog.Accepted:
            return

        result = payment_service.mark_paid(
            firm=self._firm,
            invoice_number=inv,
            status=dlg.status(),
            payment_date=dlg.payment_date(),
            notes=dlg.notes() or None,
            config=self._config,
        )

        if result.success:
            self.refresh()
            self.paymentUpdated.emit()
        else:
            show_error(self, "Payment Failed", result.message)

    # ── helpers ───────────────────────────────────────────────────

    @staticmethod
    def _parse_date(val) -> date | None:
        if val is None:
            return None
        if isinstance(val, date) and not isinstance(val, datetime):
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

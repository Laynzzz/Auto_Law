"""Tab 5 — Read-only audit log viewer with filtering and CSV export."""

import csv
from datetime import date, datetime
from pathlib import Path

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.audit_log import AUDIT_COLUMNS
from src.dataset import get_data_root

AUDIT_HEADERS = {
    "timestamp":   "Timestamp",
    "user":        "User",
    "hostname":    "Hostname",
    "action":      "Action",
    "firm":        "Firm",
    "case_key":    "Case Key",
    "field_name":  "Field",
    "old_value":   "Old Value",
    "new_value":   "New Value",
    "reason":      "Reason",
}


class AuditTab(QWidget):
    """Read-only viewer for the centralized audit_log.csv."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._firm: str | None = None
        self._config: dict | None = None
        self._all_rows: list[dict] = []

        layout = QVBoxLayout(self)

        # ── Filter bar ───────────────────────────────────────────
        bar = QHBoxLayout()

        bar.addWidget(QLabel("From:"))
        self._date_from = QDateEdit()
        self._date_from.setCalendarPopup(True)
        self._date_from.setSpecialValueText(" ")
        self._date_from.setMinimumDate(QDate(2000, 1, 1))
        self._date_from.setDate(QDate(2000, 1, 1))
        bar.addWidget(self._date_from)

        bar.addWidget(QLabel("To:"))
        self._date_to = QDateEdit()
        self._date_to.setCalendarPopup(True)
        self._date_to.setSpecialValueText(" ")
        self._date_to.setMinimumDate(QDate(2000, 1, 1))
        self._date_to.setDate(QDate(2000, 1, 1))
        bar.addWidget(self._date_to)

        bar.addWidget(QLabel("Action:"))
        self._action_combo = QComboBox()
        self._action_combo.addItem("")  # All
        bar.addWidget(self._action_combo)

        bar.addStretch()

        btn_apply = QPushButton("Apply")
        btn_apply.clicked.connect(self._apply_filters)
        bar.addWidget(btn_apply)

        btn_clear = QPushButton("Clear")
        btn_clear.clicked.connect(self._clear_filters)
        bar.addWidget(btn_clear)

        btn_refresh = QPushButton("Refresh")
        btn_refresh.clicked.connect(self.refresh)
        bar.addWidget(btn_refresh)

        btn_export = QPushButton("Export CSV...")
        btn_export.clicked.connect(self._export_csv)
        bar.addWidget(btn_export)

        layout.addLayout(bar)

        # ── Table ────────────────────────────────────────────────
        self._table = QTableWidget()
        self._table.setColumnCount(len(AUDIT_COLUMNS))
        self._table.setHorizontalHeaderLabels(
            [AUDIT_HEADERS.get(c, c) for c in AUDIT_COLUMNS]
        )
        self._table.setSortingEnabled(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        self._table.verticalHeader().setVisible(False)

        layout.addWidget(self._table)

    # ── public API ────────────────────────────────────────────────

    def set_firm(self, firm: str, config: dict | None = None):
        self._firm = firm
        self._config = config
        self._load_data()

    def refresh(self):
        self._load_data()

    # ── data loading ──────────────────────────────────────────────

    def _audit_log_path(self) -> Path:
        return get_data_root() / "invoice" / "audit_log.csv"

    def _load_data(self):
        """Read audit_log.csv into memory. Gracefully handle missing file."""
        path = self._audit_log_path()
        self._all_rows = []

        if not path.exists():
            self._populate_action_combo()
            self._apply_filters()
            return

        try:
            with open(path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    self._all_rows.append(row)
        except Exception:
            # Gracefully handle corrupt / unreadable file
            self._all_rows = []

        self._populate_action_combo()
        self._apply_filters()

    def _populate_action_combo(self):
        """Populate action combo from unique actions found in data."""
        current = self._action_combo.currentText()
        actions = sorted({r.get("action", "") for r in self._all_rows if r.get("action")})
        self._action_combo.blockSignals(True)
        self._action_combo.clear()
        self._action_combo.addItem("")  # All
        self._action_combo.addItems(actions)
        idx = self._action_combo.findText(current)
        if idx >= 0:
            self._action_combo.setCurrentIndex(idx)
        self._action_combo.blockSignals(False)

    # ── filtering ─────────────────────────────────────────────────

    def _apply_filters(self):
        rows = self._all_rows

        # Firm filter (auto-set from sidebar)
        if self._firm:
            rows = [r for r in rows if r.get("firm", "") == self._firm]

        # Date range
        qd_from = self._date_from.date()
        if qd_from != self._date_from.minimumDate():
            d_from = date(qd_from.year(), qd_from.month(), qd_from.day())
            rows = [r for r in rows if self._parse_ts_date(r.get("timestamp")) is not None
                    and self._parse_ts_date(r.get("timestamp")) >= d_from]

        qd_to = self._date_to.date()
        if qd_to != self._date_to.minimumDate():
            d_to = date(qd_to.year(), qd_to.month(), qd_to.day())
            rows = [r for r in rows if self._parse_ts_date(r.get("timestamp")) is not None
                    and self._parse_ts_date(r.get("timestamp")) <= d_to]

        # Action filter
        action = self._action_combo.currentText()
        if action:
            rows = [r for r in rows if r.get("action", "") == action]

        self._filtered_rows = rows
        self._render(rows)

    def _clear_filters(self):
        self._date_from.setDate(self._date_from.minimumDate())
        self._date_to.setDate(self._date_to.minimumDate())
        self._action_combo.setCurrentIndex(0)
        self._apply_filters()

    def _render(self, rows: list[dict]):
        self._table.setSortingEnabled(False)
        self._table.setRowCount(len(rows))

        for row_idx, row_data in enumerate(rows):
            for col_idx, col_name in enumerate(AUDIT_COLUMNS):
                text = row_data.get(col_name, "")
                item = QTableWidgetItem(str(text) if text else "")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self._table.setItem(row_idx, col_idx, item)

        self._table.setSortingEnabled(True)

    # ── export ────────────────────────────────────────────────────

    def _export_csv(self):
        rows = getattr(self, "_filtered_rows", [])
        if not rows:
            QMessageBox.information(self, "No data", "No rows to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Export Audit Log", "audit_log_export.csv",
            "CSV Files (*.csv)"
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=AUDIT_COLUMNS)
                writer.writeheader()
                for row in rows:
                    writer.writerow({col: row.get(col, "") for col in AUDIT_COLUMNS})
            QMessageBox.information(
                self, "Export Complete",
                f"Exported {len(rows)} rows to:\n{path}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", str(exc))

    # ── helpers ───────────────────────────────────────────────────

    @staticmethod
    def _parse_ts_date(val) -> date | None:
        """Parse timestamp string to date (for filtering)."""
        if not val:
            return None
        s = str(val).strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        return None

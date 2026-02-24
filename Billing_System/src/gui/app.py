"""PySide6 GUI entry-point — MainWindow with sidebar + 3 tabs."""

import sys
import traceback

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.config import load_config
from src.dataset import VALID_CASE_STATUSES, VALID_PAID_STATUSES, all_firm_names
from src.gui.cases_tab import CasesTab
from src.gui.add_case_tab import AddCaseTab
from src.gui.generate_tab import GenerateTab
from src.gui.payments_tab import PaymentsTab
from src.gui.audit_tab import AuditTab


SORTED_CASE_STATUSES = sorted(VALID_CASE_STATUSES)
SORTED_PAID_STATUSES = sorted(VALID_PAID_STATUSES)


class MainWindow(QMainWindow):
    """Application main window — sidebar + tabs."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Picerno & Associates \u2014 Billing System")
        self.resize(1400, 800)

        self._config = load_config()

        # ── central widget ────────────────────────────────────────
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)

        # ── LEFT SIDEBAR ──────────────────────────────────────────
        sidebar = QVBoxLayout()
        sidebar.setContentsMargins(4, 4, 4, 4)

        # Firm selector
        sidebar.addWidget(QLabel("Firm:"))
        self._firm_combo = QComboBox()
        self._firm_combo.setMinimumWidth(200)
        firms = all_firm_names(self._config)
        self._firm_combo.addItems(firms)
        sidebar.addWidget(self._firm_combo)

        # Filters group
        filter_box = QGroupBox("Filters")
        flay = QVBoxLayout(filter_box)

        flay.addWidget(QLabel("From:"))
        self._filter_from = QDateEdit()
        self._filter_from.setCalendarPopup(True)
        self._filter_from.setSpecialValueText(" ")  # blank when cleared
        self._filter_from.setMinimumDate(QDate(2000, 1, 1))
        self._filter_from.setDate(QDate(2000, 1, 1))  # effectively blank
        flay.addWidget(self._filter_from)

        flay.addWidget(QLabel("To:"))
        self._filter_to = QDateEdit()
        self._filter_to.setCalendarPopup(True)
        self._filter_to.setSpecialValueText(" ")
        self._filter_to.setMinimumDate(QDate(2000, 1, 1))
        self._filter_to.setDate(QDate(2000, 1, 1))
        flay.addWidget(self._filter_to)

        flay.addWidget(QLabel("Case Status:"))
        self._filter_status = QComboBox()
        self._filter_status.addItem("")  # All
        self._filter_status.addItems(SORTED_CASE_STATUSES)
        flay.addWidget(self._filter_status)

        flay.addWidget(QLabel("Paid Status:"))
        self._filter_paid = QComboBox()
        self._filter_paid.addItem("")  # All
        self._filter_paid.addItems(SORTED_PAID_STATUSES)
        flay.addWidget(self._filter_paid)

        btn_row = QHBoxLayout()
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self._apply_filters)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_filters)
        btn_row.addWidget(apply_btn)
        btn_row.addWidget(clear_btn)
        flay.addLayout(btn_row)

        sidebar.addWidget(filter_box)
        sidebar.addStretch()

        sidebar_widget = QWidget()
        sidebar_widget.setLayout(sidebar)
        sidebar_widget.setFixedWidth(220)
        root.addWidget(sidebar_widget)

        # ── TABS ──────────────────────────────────────────────────
        self._tabs = QTabWidget()

        self._cases_tab = CasesTab()
        self._add_tab = AddCaseTab()
        self._gen_tab = GenerateTab()
        self._payments_tab = PaymentsTab()
        self._audit_tab = AuditTab()

        self._tabs.addTab(self._cases_tab, "Cases")
        self._tabs.addTab(self._add_tab, "Add Case")
        self._tabs.addTab(self._gen_tab, "Generate")
        self._tabs.addTab(self._payments_tab, "Payments")
        self._tabs.addTab(self._audit_tab, "Audit Log")

        root.addWidget(self._tabs, 1)

        # ── Status bar ────────────────────────────────────────────
        self.statusBar().showMessage("Ready")

        # ── Signal wiring ─────────────────────────────────────────
        self._firm_combo.currentTextChanged.connect(self._on_firm_changed)
        self._add_tab.caseAdded.connect(self._cases_tab.refresh)
        self._cases_tab.caseSelected.connect(self._gen_tab.set_selected_case)
        self._payments_tab.paymentUpdated.connect(self._cases_tab.refresh)
        self._gen_tab.statusMessage.connect(
            lambda msg: self.statusBar().showMessage(msg, 10_000)
        )

        # ── Initial load ─────────────────────────────────────────
        if firms:
            self._on_firm_changed(firms[0])

    # ── firm change ───────────────────────────────────────────────

    def _on_firm_changed(self, firm: str):
        self._cases_tab.load_firm(firm, self._config)
        self._add_tab.set_firm(firm, self._config)
        self._gen_tab.set_firm(firm, self._config)
        self._payments_tab.set_firm(firm, self._config)
        self._audit_tab.set_firm(firm, self._config)

    # ── filters ───────────────────────────────────────────────────

    def _apply_filters(self):
        from datetime import date as _date

        filters: dict = {}

        qd_from = self._filter_from.date()
        if qd_from != self._filter_from.minimumDate():
            filters["date_from"] = _date(
                qd_from.year(), qd_from.month(), qd_from.day()
            )

        qd_to = self._filter_to.date()
        if qd_to != self._filter_to.minimumDate():
            filters["date_to"] = _date(
                qd_to.year(), qd_to.month(), qd_to.day()
            )

        cs = self._filter_status.currentText()
        if cs:
            filters["case_status"] = cs

        ps = self._filter_paid.currentText()
        if ps:
            filters["paid_status"] = ps

        self._cases_tab.apply_filters(filters)

    def _clear_filters(self):
        self._filter_from.setDate(self._filter_from.minimumDate())
        self._filter_to.setDate(self._filter_to.minimumDate())
        self._filter_status.setCurrentIndex(0)
        self._filter_paid.setCurrentIndex(0)
        self._cases_tab.apply_filters({})


# ── Global exception handler ─────────────────────────────────────

def _excepthook(exc_type, exc_val, exc_tb):
    tb = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
    QMessageBox.critical(
        None,
        "Unexpected Error",
        f"{exc_type.__name__}: {exc_val}\n\nSee details for full traceback.",
    )
    # Still print to stderr for logging
    sys.__excepthook__(exc_type, exc_val, exc_tb)


# ── Entry-point ──────────────────────────────────────────────────

def main():
    sys.excepthook = _excepthook
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

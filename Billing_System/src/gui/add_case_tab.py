"""Tab 2 — Add / Update Case form."""

from datetime import date

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.dataset import VALID_CASE_STATUSES
from src.services import case_service


SORTED_STATUSES = sorted(VALID_CASE_STATUSES)


class AddCaseTab(QWidget):
    """Form for adding or updating a case."""

    caseAdded = Signal()  # emitted after successful add/update

    def __init__(self, parent=None):
        super().__init__(parent)
        self._firm: str | None = None
        self._config: dict | None = None

        outer = QVBoxLayout(self)

        # Firm label
        self._firm_label = QLabel("No firm selected")
        self._firm_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        outer.addWidget(self._firm_label)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)

        # Appearance date
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setDate(date.today())
        self._date_edit.setDisplayFormat("yyyy-MM-dd")
        form.addRow("Appearance Date:", self._date_edit)

        # Index number
        self._index_edit = QLineEdit()
        self._index_edit.setPlaceholderText("e.g. 2026-001234")
        form.addRow("Index Number:", self._index_edit)

        # Case caption
        self._caption_edit = QLineEdit()
        self._caption_edit.setPlaceholderText("People v. Smith")
        form.addRow("Case Caption:", self._caption_edit)

        # Court
        self._court_edit = QLineEdit()
        form.addRow("Court:", self._court_edit)

        # Charge amount
        self._amount_spin = QDoubleSpinBox()
        self._amount_spin.setRange(0.0, 999_999.99)
        self._amount_spin.setDecimals(2)
        self._amount_spin.setPrefix("$ ")
        self._amount_spin.setValue(0.0)
        form.addRow("Charge Amount:", self._amount_spin)

        # Outcome
        self._outcome_edit = QLineEdit()
        form.addRow("Outcome:", self._outcome_edit)

        # Case status
        self._status_combo = QComboBox()
        self._status_combo.addItem("")  # blank default
        self._status_combo.addItems(SORTED_STATUSES)
        form.addRow("Case Status:", self._status_combo)

        # Notes
        self._notes_edit = QTextEdit()
        self._notes_edit.setMaximumHeight(80)
        form.addRow("Notes:", self._notes_edit)

        outer.addLayout(form)

        # Buttons
        btn_row = QHBoxLayout()
        self._submit_btn = QPushButton("Submit")
        self._submit_btn.clicked.connect(self._on_submit)
        self._clear_btn = QPushButton("Clear")
        self._clear_btn.clicked.connect(self._clear_form)
        btn_row.addStretch()
        btn_row.addWidget(self._submit_btn)
        btn_row.addWidget(self._clear_btn)
        outer.addLayout(btn_row)

        outer.addStretch()

    # ── public API ────────────────────────────────────────────────

    def set_firm(self, firm: str, config: dict | None = None):
        self._firm = firm
        self._config = config
        self._firm_label.setText(f"Firm: {firm}")

    # ── submit ────────────────────────────────────────────────────

    def _on_submit(self):
        if not self._firm:
            QMessageBox.warning(self, "No firm", "Select a firm first.")
            return

        index_number = self._index_edit.text().strip()
        caption = self._caption_edit.text().strip()
        amount = self._amount_spin.value()

        if not index_number or not caption:
            QMessageBox.warning(
                self, "Missing fields",
                "Index Number and Case Caption are required."
            )
            return

        appearance_date = self._date_edit.date().toString("yyyy-MM-dd")
        court = self._court_edit.text().strip() or None
        outcome = self._outcome_edit.text().strip() or None
        status = self._status_combo.currentText() or None
        notes = self._notes_edit.toPlainText().strip() or None

        result = case_service.add_or_update_case(
            firm=self._firm,
            appearance_date=appearance_date,
            index_number=index_number,
            case_caption=caption,
            charge_amount=amount,
            court=court,
            outcome=outcome,
            case_status=status,
            notes=notes,
            config=self._config,
        )

        if result.success:
            action = (result.data or {}).get("action", "saved")
            QMessageBox.information(
                self, "Success",
                f"Case {action}: {caption} ({appearance_date})"
            )
            self._clear_form()
            self.caseAdded.emit()
        else:
            QMessageBox.warning(self, "Failed", result.message)

    # ── clear ─────────────────────────────────────────────────────

    def _clear_form(self):
        self._date_edit.setDate(date.today())
        self._index_edit.clear()
        self._caption_edit.clear()
        self._court_edit.clear()
        self._amount_spin.setValue(0.0)
        self._outcome_edit.clear()
        self._status_combo.setCurrentIndex(0)
        self._notes_edit.clear()

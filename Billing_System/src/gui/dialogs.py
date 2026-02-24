"""Shared dialogs used across GUI tabs."""

from datetime import date

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QVBoxLayout,
)


# ------------------------------------------------------------------
# EditChargeDialog
# ------------------------------------------------------------------

class EditChargeDialog(QDialog):
    """Dialog for editing charge_amount on an existing case.

    If *invoice_sent* is ``True`` a warning is shown and a reason is
    mandatory before the user can click OK.
    """

    def __init__(self, current_amount: float, invoice_sent: bool = False,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Charge Amount")
        self.setMinimumWidth(360)

        layout = QVBoxLayout(self)

        # Current value
        layout.addWidget(QLabel(f"Current charge: ${current_amount:,.2f}"))

        # Warning when invoice already sent
        self._invoice_sent = invoice_sent
        if invoice_sent:
            warn = QLabel(
                "This case already has an invoice sent date.\n"
                "A reason is required for the change."
            )
            warn.setStyleSheet("color: #b00; font-weight: bold;")
            warn.setWordWrap(True)
            layout.addWidget(warn)

        # New amount
        layout.addWidget(QLabel("New amount:"))
        self._spin = QDoubleSpinBox()
        self._spin.setRange(0.0, 999_999.99)
        self._spin.setDecimals(2)
        self._spin.setPrefix("$ ")
        self._spin.setValue(current_amount)
        layout.addWidget(self._spin)

        # Reason
        layout.addWidget(QLabel("Reason for change:"))
        self._reason = QLineEdit()
        self._reason.setPlaceholderText(
            "Required" if invoice_sent else "Optional"
        )
        layout.addWidget(self._reason)

        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # --- properties ---------------------------------------------------

    def new_amount(self) -> float:
        return self._spin.value()

    def reason(self) -> str:
        return self._reason.text().strip()

    # --- internal -----------------------------------------------------

    def _on_accept(self):
        if self._invoice_sent and not self.reason():
            QMessageBox.warning(
                self, "Reason required",
                "Please provide a reason when the invoice has already been sent."
            )
            return
        self.accept()


# ------------------------------------------------------------------
# MarkPaymentDialog
# ------------------------------------------------------------------

class MarkPaymentDialog(QDialog):
    """Dialog for marking an invoice as Paid / Unpaid / Partial."""

    _STATUSES = ["Paid", "Unpaid", "Partial"]

    def __init__(self, invoice_number: str, case_caption: str,
                 current_status: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Mark Payment")
        self.setMinimumWidth(380)

        layout = QVBoxLayout(self)

        # Read-only invoice info
        info = QLabel(f"Invoice: {invoice_number} — {case_caption}")
        info.setWordWrap(True)
        layout.addWidget(info)

        # Status combo
        layout.addWidget(QLabel("Payment status:"))
        self._status_combo = QComboBox()
        self._status_combo.addItems(self._STATUSES)
        idx = self._status_combo.findText(current_status)
        if idx >= 0:
            self._status_combo.setCurrentIndex(idx)
        else:
            # Default to Unpaid if unknown
            self._status_combo.setCurrentIndex(1)
        self._status_combo.currentTextChanged.connect(self._on_status_changed)
        layout.addWidget(self._status_combo)

        # Payment date
        layout.addWidget(QLabel("Payment date:"))
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setDate(QDate.currentDate())
        self._date_edit.setDisplayFormat("yyyy-MM-dd")
        layout.addWidget(self._date_edit)

        # Notes
        layout.addWidget(QLabel("Notes (optional):"))
        self._notes = QLineEdit()
        layout.addWidget(self._notes)

        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        # Apply initial enable/disable state
        self._on_status_changed(self._status_combo.currentText())

    # --- properties ---------------------------------------------------

    def status(self) -> str:
        return self._status_combo.currentText()

    def payment_date(self) -> str | None:
        if self._status_combo.currentText() == "Unpaid":
            return None
        qd = self._date_edit.date()
        return date(qd.year(), qd.month(), qd.day()).isoformat()

    def notes(self) -> str:
        return self._notes.text().strip()

    # --- internal -----------------------------------------------------

    def _on_status_changed(self, status: str):
        self._date_edit.setEnabled(status != "Unpaid")

    def _on_accept(self):
        if self._status_combo.currentText() == "Paid" and not self._date_edit.isEnabled():
            QMessageBox.warning(
                self, "Date required",
                "Please provide a payment date when marking as Paid."
            )
            return
        self.accept()


# ------------------------------------------------------------------
# show_error helper
# ------------------------------------------------------------------

def show_error(parent, title: str, message: str, details: str | None = None):
    """Show a critical error dialog with optional copyable details."""
    box = QMessageBox(parent)
    box.setIcon(QMessageBox.Critical)
    box.setWindowTitle(title)
    box.setText(message)
    if details:
        box.setDetailedText(details)
    box.exec()

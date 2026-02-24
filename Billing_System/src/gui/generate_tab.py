"""Tab 3 — Generate PDFs and email drafts."""

from datetime import date

from PySide6.QtCore import Qt, QThreadPool, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from src.services import doc_service, email_service
from src.gui.dialogs import show_error
from src.gui.workers import ServiceWorker


class GenerateTab(QWidget):
    """Generate PDFs and create Outlook email drafts."""

    statusMessage = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._firm: str | None = None
        self._config: dict | None = None
        self._selected_case: dict | None = None
        self._busy = False
        self._pool = QThreadPool.globalInstance()

        layout = QVBoxLayout(self)

        # ── Daily section ─────────────────────────────────────────
        daily_box = QGroupBox("Daily (Selected Case)")
        daily_lay = QHBoxLayout(daily_box)
        self._daily_info = QLabel("No case selected")
        daily_lay.addWidget(self._daily_info, 1)
        self._btn_daily_pdf = QPushButton("Generate Daily PDF")
        self._btn_daily_pdf.clicked.connect(self._gen_daily_pdf)
        daily_lay.addWidget(self._btn_daily_pdf)
        self._btn_daily_email = QPushButton("Draft Daily Email")
        self._btn_daily_email.clicked.connect(self._draft_daily_email)
        daily_lay.addWidget(self._btn_daily_email)
        layout.addWidget(daily_box)

        # ── Weekly section ────────────────────────────────────────
        weekly_box = QGroupBox("Weekly Statement")
        weekly_lay = QHBoxLayout(weekly_box)
        weekly_lay.addWidget(QLabel("Week of:"))
        self._weekly_date = QDateEdit()
        self._weekly_date.setCalendarPopup(True)
        self._weekly_date.setDate(date.today())
        self._weekly_date.setDisplayFormat("yyyy-MM-dd")
        weekly_lay.addWidget(self._weekly_date)
        self._btn_weekly_pdf = QPushButton("Generate Weekly PDF")
        self._btn_weekly_pdf.clicked.connect(self._gen_weekly_pdf)
        weekly_lay.addWidget(self._btn_weekly_pdf)
        self._btn_weekly_email = QPushButton("Draft Weekly Email")
        self._btn_weekly_email.clicked.connect(self._draft_weekly_email)
        weekly_lay.addWidget(self._btn_weekly_email)
        layout.addWidget(weekly_box)

        # ── Monthly section ───────────────────────────────────────
        monthly_box = QGroupBox("Monthly Statement")
        monthly_lay = QHBoxLayout(monthly_box)
        monthly_lay.addWidget(QLabel("Year:"))
        self._month_year = QSpinBox()
        self._month_year.setRange(2020, 2099)
        self._month_year.setValue(date.today().year)
        monthly_lay.addWidget(self._month_year)
        monthly_lay.addWidget(QLabel("Month:"))
        self._month_month = QComboBox()
        for i in range(1, 13):
            self._month_month.addItem(
                date(2000, i, 1).strftime("%B"), i
            )
        self._month_month.setCurrentIndex(date.today().month - 1)
        monthly_lay.addWidget(self._month_month)
        self._btn_monthly_pdf = QPushButton("Generate Monthly PDF")
        self._btn_monthly_pdf.clicked.connect(self._gen_monthly_pdf)
        monthly_lay.addWidget(self._btn_monthly_pdf)
        self._btn_monthly_email = QPushButton("Draft Monthly Email")
        self._btn_monthly_email.clicked.connect(self._draft_monthly_email)
        monthly_lay.addWidget(self._btn_monthly_email)
        layout.addWidget(monthly_box)

        # ── Ledger section ────────────────────────────────────────
        ledger_box = QGroupBox("Ledger Export")
        ledger_lay = QHBoxLayout(ledger_box)
        ledger_lay.addWidget(QLabel("As of:"))
        self._ledger_date = QDateEdit()
        self._ledger_date.setCalendarPopup(True)
        self._ledger_date.setDate(date.today())
        self._ledger_date.setDisplayFormat("yyyy-MM-dd")
        ledger_lay.addWidget(self._ledger_date)
        self._btn_ledger = QPushButton("Export Ledger")
        self._btn_ledger.clicked.connect(self._export_ledger)
        ledger_lay.addWidget(self._btn_ledger)
        layout.addWidget(ledger_box)

        # ── Progress label ────────────────────────────────────────
        self._progress = QLabel("")
        self._progress.setWordWrap(True)
        layout.addWidget(self._progress)

        layout.addStretch()

        self._all_buttons = [
            self._btn_daily_pdf, self._btn_daily_email,
            self._btn_weekly_pdf, self._btn_weekly_email,
            self._btn_monthly_pdf, self._btn_monthly_email,
            self._btn_ledger,
        ]
        self._update_daily_buttons()

    # ── public API ────────────────────────────────────────────────

    def set_firm(self, firm: str, config: dict | None = None):
        self._firm = firm
        self._config = config

    def set_selected_case(self, case: dict | None):
        self._selected_case = case
        if case:
            caption = case.get("case_caption", "")
            ad = case.get("appearance_date", "")
            self._daily_info.setText(f"{caption}  ({ad})")
        else:
            self._daily_info.setText("No case selected")
        self._update_daily_buttons()

    # ── button state ──────────────────────────────────────────────

    def _update_daily_buttons(self):
        has_case = self._selected_case is not None
        self._btn_daily_pdf.setEnabled(has_case and not self._busy)
        self._btn_daily_email.setEnabled(has_case and not self._busy)

    def _set_busy(self, busy: bool):
        self._busy = busy
        for btn in self._all_buttons:
            btn.setEnabled(not busy)
        if not busy:
            self._update_daily_buttons()

    # ── worker helpers ────────────────────────────────────────────

    def _run(self, fn, *args, **kwargs):
        self._set_busy(True)
        self._progress.setText("Working...")
        self.statusMessage.emit("Working...")
        worker = ServiceWorker(fn, *args, **kwargs)
        worker.signals.finished.connect(self._on_finished)
        worker.signals.error.connect(self._on_error)
        self._pool.start(worker)

    def _on_finished(self, result):
        self._set_busy(False)
        if result.success:
            msg = result.message
            # Append path info if available
            data = result.data or {}
            for key in ("pdf_path", "xlsx_path"):
                if data.get(key):
                    msg += f"\n{key}: {data[key]}"
            self._progress.setText(f"Done: {msg}")
            self.statusMessage.emit(f"Done: {result.message}")
        else:
            self._progress.setText(f"Failed: {result.message}")
            self.statusMessage.emit(f"Failed: {result.message}")

    def _on_error(self, err_tuple):
        self._set_busy(False)
        exc_type, exc_val, tb_str = err_tuple
        self._progress.setText(f"Error: {exc_val}")
        self.statusMessage.emit(f"Error: {exc_val}")
        show_error(self, "Worker Error", str(exc_val), details=tb_str)

    # ── case helpers ──────────────────────────────────────────────

    def _case_index(self) -> str:
        return str(self._selected_case["index_number"])

    def _case_date(self) -> str:
        ad = self._selected_case["appearance_date"]
        if isinstance(ad, date):
            return ad.isoformat()
        return str(ad).strip()

    # ── action handlers ───────────────────────────────────────────

    def _gen_daily_pdf(self):
        if not self._firm or not self._selected_case:
            return
        self._run(
            doc_service.generate_daily,
            firm=self._firm,
            index_number=self._case_index(),
            appearance_date=self._case_date(),
            config=self._config,
        )

    def _draft_daily_email(self):
        if not self._firm or not self._selected_case:
            return
        self._run(
            email_service.draft_daily,
            firm=self._firm,
            index_number=self._case_index(),
            appearance_date=self._case_date(),
            config=self._config,
        )

    def _gen_weekly_pdf(self):
        if not self._firm:
            return
        self._run(
            doc_service.generate_weekly,
            firm=self._firm,
            week_of=self._weekly_date.date().toString("yyyy-MM-dd"),
            config=self._config,
        )

    def _draft_weekly_email(self):
        if not self._firm:
            return
        self._run(
            email_service.draft_weekly,
            firm=self._firm,
            week_of=self._weekly_date.date().toString("yyyy-MM-dd"),
            config=self._config,
        )

    def _gen_monthly_pdf(self):
        if not self._firm:
            return
        self._run(
            doc_service.generate_monthly,
            firm=self._firm,
            year=self._month_year.value(),
            month=self._month_month.currentData(),
            config=self._config,
        )

    def _draft_monthly_email(self):
        if not self._firm:
            return
        self._run(
            email_service.draft_monthly,
            firm=self._firm,
            year=self._month_year.value(),
            month=self._month_month.currentData(),
            config=self._config,
        )

    def _export_ledger(self):
        if not self._firm:
            return
        self._run(
            doc_service.export_ledger,
            firm=self._firm,
            as_of=self._ledger_date.date().toString("yyyy-MM-dd"),
            config=self._config,
        )

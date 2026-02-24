"""Email-draft service — create Outlook drafts for invoices and statements."""

from __future__ import annotations

import calendar
from datetime import date as _date
from pathlib import Path

from src.config import get_firm, load_config
from src.dataset import (
    all_firm_names,
    find_row_by_key,
    get_data_root,
    week_range,
)
from src.email_draft import create_draft
from src.services import ServiceResult


# ── Helpers ──────────────────────────────────────────────────────────


def _resolve_config(config: dict | None) -> dict:
    if config is None:
        return load_config()
    return config


def _validate_firm(firm: str, config: dict) -> str | None:
    known = all_firm_names(config)
    if firm not in known:
        return f"Firm '{firm}' not found. Available: {known}"
    return None


def _parse_date(val) -> _date:
    """Coerce a value to a date object."""
    from datetime import datetime
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, _date):
        return val
    return _date.fromisoformat(str(val).split(" ")[0])


def _ordinal(n: int) -> str:
    """Return day with ordinal suffix (1st, 2nd, 3rd, etc.)."""
    if 11 <= n % 100 <= 13:
        return f"{n}th"
    return f"{n}{('th','st','nd','rd')[min(n % 10, 4)] if n % 10 < 4 else 'th'}"


def _format_date_long(d: _date) -> str:
    """Format date as 'February 20th, 2026'."""
    return f"{calendar.month_name[d.month]} {_ordinal(d.day)}, {d.year}"


# ── Path resolution (mirrors doc generation output paths) ────────────


def _daily_pdf_path(firm: str, case: dict) -> Path:
    """Resolve the expected PDF path for a daily per-diem invoice."""
    dt = _parse_date(case["appearance_date"])
    date_prefix = dt.strftime("%m-%d-%Y")
    caption = str(case.get("case_caption") or "case")
    monday, _ = week_range(dt)
    week_folder = f"Week of {monday.strftime('%m-%d-%Y')}"

    return (
        get_data_root() / "invoice" / firm
        / str(dt.year) / dt.strftime("%b") / week_folder
        / "report" / "pdf" / f"{date_prefix} {caption}.pdf"
    )


def _weekly_pdf_path(firm: str, monday: _date) -> Path:
    """Resolve the expected PDF path for a weekly statement."""
    date_prefix = monday.strftime("%m-%d-%Y")
    week_folder = f"Week of {date_prefix}"

    return (
        get_data_root() / "invoice" / firm
        / str(monday.year) / monday.strftime("%b") / week_folder
        / f"{week_folder}.pdf"
    )


def _monthly_pdf_path(firm: str, year: int, month: int) -> Path:
    """Resolve the expected PDF path for a monthly statement."""
    month_abbr = _date(year, month, 1).strftime("%b")
    filename = f"Monthly Statement {month_abbr} {year}"

    return (
        get_data_root() / "invoice" / firm
        / str(year) / month_abbr / f"{filename}.pdf"
    )


# ── Email body templates (HTML) ──────────────────────────────────────


def _daily_body(firm: str, case: dict) -> str:
    dt = _parse_date(case["appearance_date"])
    caption = case.get("case_caption", "")
    index = case.get("index_number", "")
    inv = case.get("invoice_number", "")
    amount = case.get("charge_amount", 0)

    return (
        f"<p>Dear Counsel,</p>"
        f"<p>Please find attached the per diem invoice for the following appearance:</p>"
        f"<table style='border-collapse:collapse; margin:10px 0;'>"
        f"<tr><td style='padding:2px 12px 2px 0;'><b>Date:</b></td>"
        f"<td>{_format_date_long(dt)}</td></tr>"
        f"<tr><td style='padding:2px 12px 2px 0;'><b>Case:</b></td>"
        f"<td>{caption}</td></tr>"
        f"<tr><td style='padding:2px 12px 2px 0;'><b>Index #:</b></td>"
        f"<td>{index}</td></tr>"
        f"<tr><td style='padding:2px 12px 2px 0;'><b>Invoice #:</b></td>"
        f"<td>{inv}</td></tr>"
        f"<tr><td style='padding:2px 12px 2px 0;'><b>Amount:</b></td>"
        f"<td>${float(amount):,.2f}</td></tr>"
        f"</table>"
        f"<p>Thank you.</p>"
        f"<br>"
    )


def _weekly_body(firm: str, monday: _date, friday: _date) -> str:
    return (
        f"<p>Dear Counsel,</p>"
        f"<p>Please find attached the weekly statement of account for "
        f"<b>{firm}</b> covering the period "
        f"{_format_date_long(monday)} &ndash; {_format_date_long(friday)}.</p>"
        f"<p>This statement summarizes invoices previously sent. "
        f"No new charges are added.</p>"
        f"<p>Thank you.</p>"
        f"<br>"
    )


def _monthly_body(firm: str, year: int, month: int) -> str:
    month_name = calendar.month_name[month]
    return (
        f"<p>Dear Counsel,</p>"
        f"<p>Please find attached the monthly statement of account for "
        f"<b>{firm}</b> for <b>{month_name} {year}</b>.</p>"
        f"<p>This statement summarizes invoices previously sent. "
        f"No new charges are added.</p>"
        f"<p>Thank you.</p>"
        f"<br>"
    )


# ── Public API ───────────────────────────────────────────────────────


def draft_daily(
    firm: str,
    index_number: str,
    appearance_date: str,
    config: dict | None = None,
) -> ServiceResult:
    """Create an Outlook draft for a daily per-diem invoice.

    The PDF must already exist (run ``generate-daily`` first).
    Returns draft metadata in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    case = find_row_by_key(firm, index_number, appearance_date)
    if case is None:
        return ServiceResult(
            success=False,
            message=f"Case not found: index={index_number}, date={appearance_date} in firm '{firm}'.",
        )

    inv_num = case.get("invoice_number")
    if not inv_num:
        return ServiceResult(
            success=False,
            message="Case has no invoice number. Run 'assign-invoices' first.",
        )

    pdf = _daily_pdf_path(firm, case)
    if not pdf.exists():
        return ServiceResult(
            success=False,
            message=f"Invoice PDF not found: {pdf}\nRun 'generate-daily' first.",
        )

    firm_cfg = get_firm(firm, config)
    to = firm_cfg.get("billing_email", "")
    if not to:
        return ServiceResult(
            success=False,
            message=f"No billing_email configured for firm '{firm}'. Update config.json.",
        )

    cc_list = firm_cfg.get("cc_emails", [])
    cc = "; ".join(cc_list) if cc_list else None

    caption = case.get("case_caption", "")
    subject = f"Per Diem Invoice {inv_num} - {caption}"
    body = _daily_body(firm, case)

    try:
        meta = create_draft(to=to, subject=subject, body_html=body, cc=cc, attachment_paths=[pdf])
    except (OSError, FileNotFoundError) as exc:
        return ServiceResult(success=False, message=str(exc))

    return ServiceResult(
        success=True,
        message=f"Draft created in Outlook:\n  Subject: {subject}\n  To: {to}\n  Attachments: {meta['attachments']}",
        data=meta,
    )


def draft_weekly(
    firm: str,
    week_of: str,
    config: dict | None = None,
) -> ServiceResult:
    """Create an Outlook draft for a weekly statement.

    The PDF must already exist (run ``generate-weekly`` first).
    Returns draft metadata in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    try:
        ref = _date.fromisoformat(week_of)
    except ValueError:
        return ServiceResult(success=False, message=f"Invalid date: {week_of}. Use YYYY-MM-DD.")

    monday, friday = week_range(ref)
    pdf = _weekly_pdf_path(firm, monday)

    if not pdf.exists():
        return ServiceResult(
            success=False,
            message=f"Weekly statement PDF not found: {pdf}\nRun 'generate-weekly' first.",
        )

    firm_cfg = get_firm(firm, config)
    to = firm_cfg.get("billing_email", "")
    if not to:
        return ServiceResult(
            success=False,
            message=f"No billing_email configured for firm '{firm}'. Update config.json.",
        )

    cc_list = firm_cfg.get("cc_emails", [])
    cc = "; ".join(cc_list) if cc_list else None

    week_label = monday.strftime("%m/%d/%Y")
    subject = f"Weekly Statement of Account - {firm} - Week of {week_label}"
    body = _weekly_body(firm, monday, friday)

    try:
        meta = create_draft(to=to, subject=subject, body_html=body, cc=cc, attachment_paths=[pdf])
    except (OSError, FileNotFoundError) as exc:
        return ServiceResult(success=False, message=str(exc))

    return ServiceResult(
        success=True,
        message=f"Draft created in Outlook:\n  Subject: {subject}\n  To: {to}\n  Attachments: {meta['attachments']}",
        data=meta,
    )


def draft_monthly(
    firm: str,
    year: int,
    month: int,
    config: dict | None = None,
) -> ServiceResult:
    """Create an Outlook draft for a monthly statement.

    The PDF must already exist (run ``generate-monthly`` first).
    Returns draft metadata in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    if not (1 <= month <= 12):
        return ServiceResult(success=False, message=f"Invalid month: {month}. Must be 1-12.")

    pdf = _monthly_pdf_path(firm, year, month)

    if not pdf.exists():
        return ServiceResult(
            success=False,
            message=f"Monthly statement PDF not found: {pdf}\nRun 'generate-monthly' first.",
        )

    firm_cfg = get_firm(firm, config)
    to = firm_cfg.get("billing_email", "")
    if not to:
        return ServiceResult(
            success=False,
            message=f"No billing_email configured for firm '{firm}'. Update config.json.",
        )

    cc_list = firm_cfg.get("cc_emails", [])
    cc = "; ".join(cc_list) if cc_list else None

    month_name = calendar.month_name[month]
    subject = f"Monthly Statement of Account - {firm} - {month_name} {year}"
    body = _monthly_body(firm, year, month)

    try:
        meta = create_draft(to=to, subject=subject, body_html=body, cc=cc, attachment_paths=[pdf])
    except (OSError, FileNotFoundError) as exc:
        return ServiceResult(success=False, message=str(exc))

    return ServiceResult(
        success=True,
        message=f"Draft created in Outlook:\n  Subject: {subject}\n  To: {to}\n  Attachments: {meta['attachments']}",
        data=meta,
    )

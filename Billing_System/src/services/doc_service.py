"""Document-generation service — daily invoice, weekly/monthly statements, ledger."""

from __future__ import annotations

from datetime import date as _date

from src.config import load_config
from src.dataset import all_firm_names
from src.doc_generator import generate_invoice
from src.ledger_export import export_ledger as _export_ledger
from src.monthly_statement import generate_monthly_statement
from src.services import ServiceResult
from src.weekly_statement import generate_weekly_statement


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


# ── Public API ───────────────────────────────────────────────────────


def generate_daily(
    firm: str,
    index_number: str,
    appearance_date: str,
    keep_docx: bool = False,
    config: dict | None = None,
) -> ServiceResult:
    """Generate a per-diem invoice PDF for a specific case.

    Returns the PDF path in ``data["pdf_path"]``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    try:
        pdf_path = generate_invoice(
            firm, index_number, appearance_date, config, keep_docx=keep_docx
        )
    except (ValueError, FileNotFoundError) as exc:
        return ServiceResult(success=False, message=str(exc))

    return ServiceResult(
        success=True,
        message=f"Invoice generated: {pdf_path}",
        data={"pdf_path": pdf_path},
    )


def generate_weekly(
    firm: str,
    week_of: str,
    keep_docx: bool = False,
    config: dict | None = None,
) -> ServiceResult:
    """Generate a weekly statement of account for a firm.

    *week_of* is a YYYY-MM-DD string; the Mon-Fri range is computed.
    Returns the PDF path and date range in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    try:
        ref = _date.fromisoformat(week_of)
    except ValueError:
        return ServiceResult(
            success=False,
            message=f"Invalid date: {week_of}. Use YYYY-MM-DD.",
        )

    try:
        pdf_path = generate_weekly_statement(firm, ref, config, keep_docx=keep_docx)
    except FileNotFoundError as exc:
        return ServiceResult(success=False, message=str(exc))

    from src.dataset import week_range
    monday, friday = week_range(ref)

    return ServiceResult(
        success=True,
        message=f"Weekly statement generated: {pdf_path}",
        data={"pdf_path": pdf_path, "monday": monday, "friday": friday},
    )


def generate_monthly(
    firm: str,
    year: int,
    month: int,
    keep_docx: bool = False,
    config: dict | None = None,
) -> ServiceResult:
    """Generate a monthly statement of account for a firm.

    Returns the PDF path in ``data["pdf_path"]``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    if not (1 <= month <= 12):
        return ServiceResult(
            success=False,
            message=f"Invalid month: {month}. Must be 1-12.",
        )

    try:
        pdf_path = generate_monthly_statement(
            firm, year, month, config, keep_docx=keep_docx
        )
    except FileNotFoundError as exc:
        return ServiceResult(success=False, message=str(exc))

    return ServiceResult(
        success=True,
        message=f"Monthly statement generated: {pdf_path}",
        data={"pdf_path": pdf_path},
    )


def export_ledger(
    firm: str,
    as_of: str | None = None,
    xlsx: bool = True,
    keep_docx: bool = False,
    config: dict | None = None,
) -> ServiceResult:
    """Export a firm's full-history master ledger (PDF + optional XLSX).

    Returns file paths in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    as_of_date = None
    if as_of:
        try:
            as_of_date = _date.fromisoformat(as_of)
        except ValueError:
            return ServiceResult(
                success=False,
                message=f"Invalid date: {as_of}. Use YYYY-MM-DD.",
            )

    try:
        result = _export_ledger(
            firm, as_of=as_of_date, config=config,
            keep_docx=keep_docx, xlsx=xlsx,
        )
    except FileNotFoundError as exc:
        return ServiceResult(success=False, message=str(exc))

    lines = [f"Ledger PDF: {result['pdf']}"]
    data: dict = {"pdf_path": result["pdf"]}
    if "xlsx" in result:
        lines.append(f"Ledger XLSX: {result['xlsx']}")
        data["xlsx_path"] = result["xlsx"]
    else:
        data["xlsx_path"] = None

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data=data,
    )

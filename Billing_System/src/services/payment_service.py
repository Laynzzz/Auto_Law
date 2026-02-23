"""Payment service — mark invoices as Paid, Unpaid, or Partial."""

from __future__ import annotations

from datetime import date as _date

from src.config import load_config
from src.dataset import all_firm_names
from src.payment import mark_payment
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


# ── Public API ───────────────────────────────────────────────────────


def mark_paid(
    firm: str,
    invoice_number: str,
    status: str,
    payment_date: str | None = None,
    notes: str | None = None,
    config: dict | None = None,
) -> ServiceResult:
    """Mark an invoice as Paid, Unpaid, or Partial.

    Returns the updated row in ``data["updated_row"]``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    if payment_date:
        try:
            _date.fromisoformat(payment_date)
        except ValueError:
            return ServiceResult(
                success=False,
                message=f"Invalid date: {payment_date}. Use YYYY-MM-DD.",
            )

    try:
        updated = mark_payment(firm, invoice_number, status, payment_date, notes)
    except (ValueError, FileNotFoundError) as exc:
        return ServiceResult(success=False, message=str(exc))

    lines = [f"Payment updated: {firm}"]
    lines.append(f"  Invoice:      {invoice_number}")
    lines.append(f"  Case:         {updated.get('case_caption', '')}")
    lines.append(f"  Amount:       ${float(updated.get('charge_amount') or 0):,.2f}")
    lines.append(f"  Status:       {updated.get('paid_status', '')}")
    lines.append(f"  Payment date: {updated.get('payment_date', '')}")
    if notes:
        lines.append(f"  Notes:        {notes}")

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={"updated_row": updated, "invoice_number": invoice_number},
    )

"""CLI entrypoint for the law firm billing system."""

import click

from src.config import load_config
from src.dataset import (
    COLUMNS,
    VALID_CASE_STATUSES,
    all_firm_names,
    create_all_workbooks,
    create_workbook,
    dataset_path,
    load_dataset,
    upsert_row,
    validate_dataset,
)
from src.doc_generator import generate_invoice
from src.invoice_number import assign_invoice_numbers
from src.ledger_export import export_ledger
from src.legacy_import import import_legacy_invoice
from src.monthly_statement import generate_monthly_statement
from src.payment import mark_payment, find_by_invoice_number, VALID_STATUSES
from src.weekly_statement import generate_weekly_statement


@click.group()
@click.pass_context
def cli(ctx):
    """Law Firm Billing System - manage cases, invoices, and payments."""
    ctx.ensure_object(dict)
    ctx.obj["config"] = load_config()


# ── Phase 1: working command ──────────────────────────────────────────


@cli.command("config-check")
@click.pass_context
def config_check(ctx):
    """Load config and print a summary."""
    cfg = ctx.obj["config"]

    click.echo("=== Config Check ===\n")

    click.echo("Firms:")
    for firm in cfg["firms"]:
        click.echo(f"  - {firm['name']} ({firm['initials']})")

    click.echo(f"\nPaths:")
    for label, path in cfg["paths"].items():
        click.echo(f"  {label}: {path}")

    numbering = cfg.get("invoice_numbering", {})
    click.echo(f"\nInvoice numbering:")
    click.echo(f"  format: {numbering.get('format', 'N/A')}")
    click.echo(f"  yearly_reset: {numbering.get('yearly_reset', 'N/A')}")

    click.echo("\nConfig OK.")


# ── Phase 2: dataset commands ─────────────────────────────────────────


@cli.command("init-dataset")
@click.option("--firm", default=None, help="Firm name (omit to init all firms).")
@click.option("--force", is_flag=True, help="Overwrite existing file (erases all data).")
@click.pass_context
def init_dataset(ctx, firm, force):
    """Create master_cases.xlsx per firm (one file per law firm)."""
    cfg = ctx.obj["config"]
    firms = [firm] if firm else all_firm_names(cfg)

    for name in firms:
        try:
            path = create_workbook(name, overwrite=force)
            click.echo(f"Created: {path}")
        except FileExistsError as e:
            raise click.ClickException(str(e))

    click.echo(f"\n  Sheet: cases")
    click.echo(f"  Columns ({len(COLUMNS)}): {', '.join(COLUMNS)}")
    click.echo(f"  Initialized {len(firms)} firm(s).")


@cli.command("validate-dataset")
@click.option("--firm", default=None, help="Firm name (omit to validate all firms).")
@click.pass_context
def validate_dataset_cmd(ctx, firm):
    """Validate master_cases.xlsx for each firm."""
    cfg = ctx.obj["config"]
    firms = [firm] if firm else all_firm_names(cfg)
    total_errors = 0

    for name in firms:
        path = dataset_path(name)
        click.echo(f"--- {name} ---")
        click.echo(f"  File: {path}")

        errors = validate_dataset(name)
        if errors:
            click.echo(f"  FAILED - {len(errors)} error(s):")
            for err in errors:
                click.echo(f"    - {err}")
            total_errors += len(errors)
        else:
            rows = load_dataset(name)
            click.echo(f"  OK - {len(rows)} data row(s)")
        click.echo()

    if total_errors:
        click.echo(f"Total errors across all firms: {total_errors}")
        raise SystemExit(1)
    else:
        click.echo(f"All {len(firms)} firm(s) validated OK.")


# ── Phase 3: add/update case ──────────────────────────────────────────


@cli.command("add-case")
@click.option("--firm", required=True, help="Law firm name (must match config).")
@click.option("--date", "appearance_date", required=True, help="Appearance date (YYYY-MM-DD).")
@click.option("--index", "index_number", required=True, help="Case index number.")
@click.option("--caption", "case_caption", required=True, help="Case caption.")
@click.option("--amount", "charge_amount", required=True, type=float, help="Charge amount (USD).")
@click.option("--court", default=None, help="Court name.")
@click.option("--outcome", default=None, help="Case outcome.")
@click.option("--status", "case_status", default=None, help=f"Case status: {', '.join(sorted(VALID_CASE_STATUSES))}.")
@click.option("--notes", default=None, help="Additional notes.")
@click.pass_context
def add_case(ctx, firm, appearance_date, index_number, case_caption, charge_amount,
             court, outcome, case_status, notes):
    """Add or update a case in a firm's dataset."""
    cfg = ctx.obj["config"]

    # Validate firm exists in config
    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    row_data = {
        "appearance_date": appearance_date,
        "index_number": index_number,
        "case_caption": case_caption,
        "charge_amount": charge_amount,
        "court": court,
        "outcome": outcome,
        "case_status": case_status,
        "notes": notes,
    }

    # Remove None values so updates don't blank out existing fields
    row_data = {k: v for k, v in row_data.items() if v is not None}

    try:
        action = upsert_row(firm, row_data)
    except FileNotFoundError as e:
        raise click.ClickException(str(e))

    click.echo(f"Case {action}: {firm}")
    click.echo(f"  index_number:    {index_number}")
    click.echo(f"  appearance_date: {appearance_date}")
    click.echo(f"  case_caption:    {case_caption}")
    click.echo(f"  charge_amount:   {charge_amount}")
    if court:
        click.echo(f"  court:           {court}")
    if outcome:
        click.echo(f"  outcome:         {outcome}")
    if case_status:
        click.echo(f"  case_status:     {case_status}")
    if notes:
        click.echo(f"  notes:           {notes}")


# ── Phase 4: invoice numbering ────────────────────────────────────────


@cli.command("assign-invoices")
@click.option("--firm", default=None, help="Firm name (omit to assign for all firms).")
@click.pass_context
def assign_invoices(ctx, firm):
    """Assign invoice numbers to cases that don't have one yet."""
    cfg = ctx.obj["config"]
    firms = [firm] if firm else all_firm_names(cfg)

    for name in firms:
        click.echo(f"--- {name} ---")
        try:
            assigned = assign_invoice_numbers(name, cfg)
        except FileNotFoundError as e:
            raise click.ClickException(str(e))

        if assigned:
            for inv in assigned:
                click.echo(f"  Assigned: {inv}")
            click.echo(f"  Total new: {len(assigned)}")
        else:
            click.echo("  No cases need invoice numbers.")
        click.echo()


# ── Phase 5: generate daily invoice ───────────────────────────────────


@cli.command("generate-daily")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--index", "index_number", required=True, help="Case index number.")
@click.option("--date", "appearance_date", required=True, help="Appearance date (YYYY-MM-DD).")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def generate_daily(ctx, firm, index_number, appearance_date, keep_docx):
    """Generate a per-diem invoice PDF for a specific case."""
    cfg = ctx.obj["config"]

    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    try:
        pdf_path = generate_invoice(
            firm, index_number, appearance_date, cfg, keep_docx=keep_docx
        )
    except (ValueError, FileNotFoundError) as e:
        raise click.ClickException(str(e))

    click.echo(f"Invoice generated: {pdf_path}")


# ── Phase 7: weekly statement ─────────────────────────────────────────


@cli.command("generate-weekly")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--week-of", required=True, help="Any date within the week (YYYY-MM-DD). Mon-Fri range is computed.")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def generate_weekly(ctx, firm, week_of, keep_docx):
    """Generate a weekly statement of account for a firm."""
    from datetime import date as _date

    cfg = ctx.obj["config"]
    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    try:
        ref = _date.fromisoformat(week_of)
    except ValueError:
        raise click.ClickException(f"Invalid date: {week_of}. Use YYYY-MM-DD.")

    try:
        pdf_path = generate_weekly_statement(firm, ref, cfg, keep_docx=keep_docx)
    except FileNotFoundError as e:
        raise click.ClickException(str(e))

    click.echo(f"Weekly statement generated: {pdf_path}")


# ── Phase 8: monthly statement ───────────────────────────────────────


@cli.command("generate-monthly")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--year", required=True, type=int, help="Year (e.g. 2026).")
@click.option("--month", required=True, type=int, help="Month number (1-12).")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def generate_monthly(ctx, firm, year, month, keep_docx):
    """Generate a monthly statement of account for a firm."""
    cfg = ctx.obj["config"]
    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    if not (1 <= month <= 12):
        raise click.ClickException(f"Invalid month: {month}. Must be 1-12.")

    try:
        pdf_path = generate_monthly_statement(firm, year, month, cfg, keep_docx=keep_docx)
    except FileNotFoundError as e:
        raise click.ClickException(str(e))

    click.echo(f"Monthly statement generated: {pdf_path}")


# ── Phase 9: master ledger export ────────────────────────────────────


@cli.command("export-ledger")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--asof", default=None, help="As-of date (YYYY-MM-DD). Defaults to today.")
@click.option("--no-xlsx", is_flag=True, help="Skip XLSX generation (PDF only).")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def export_ledger_cmd(ctx, firm, asof, no_xlsx, keep_docx):
    """Export a firm's full-history master ledger (PDF + XLSX)."""
    from datetime import date as _date

    cfg = ctx.obj["config"]
    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    as_of_date = None
    if asof:
        try:
            as_of_date = _date.fromisoformat(asof)
        except ValueError:
            raise click.ClickException(f"Invalid date: {asof}. Use YYYY-MM-DD.")

    try:
        result = export_ledger(
            firm, as_of=as_of_date, config=cfg,
            keep_docx=keep_docx, xlsx=not no_xlsx,
        )
    except FileNotFoundError as e:
        raise click.ClickException(str(e))

    click.echo(f"Ledger PDF: {result['pdf']}")
    if "xlsx" in result:
        click.echo(f"Ledger XLSX: {result['xlsx']}")


# ── Phase 10: legacy import ──────────────────────────────────────────


@cli.command("import-legacy")
@click.option("--firm", required=True, help="Law firm name (must match config).")
@click.option("--file", "file_path", required=True, type=click.Path(exists=True), help="Path to legacy monthly invoice .docx file.")
@click.pass_context
def import_legacy(ctx, firm, file_path):
    """Import cases from a legacy monthly invoice .docx into a firm's dataset."""
    cfg = ctx.obj["config"]
    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    try:
        results = import_legacy_invoice(firm, file_path)
    except (FileNotFoundError, ValueError) as e:
        raise click.ClickException(str(e))

    if not results:
        click.echo("No cases found in the file.")
        return

    inserted = 0
    updated = 0
    for action, case in results:
        label = "NEW" if action == "inserted" else "UPD"
        click.echo(
            f"  [{label}] {case['appearance_date']} | "
            f"{case['index_number']} | {case['case_caption']} | "
            f"${case['charge_amount']:.2f}"
        )
        if action == "inserted":
            inserted += 1
        else:
            updated += 1

    click.echo(f"\nImported {len(results)} case(s): {inserted} new, {updated} updated.")


# ── Phase 11: payment update ─────────────────────────────────────────


@cli.command("mark-paid")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--invoice", "invoice_number", required=True, help="Invoice number (e.g. AL2026001).")
@click.option("--status", required=True, type=click.Choice(sorted(VALID_STATUSES), case_sensitive=False),
              help="Payment status.")
@click.option("--date", "payment_date", default=None, help="Payment date (YYYY-MM-DD). Auto-set to today if marking Paid.")
@click.option("--notes", default=None, help="Optional notes about the payment.")
@click.pass_context
def mark_paid(ctx, firm, invoice_number, status, payment_date, notes):
    """Mark an invoice as Paid, Unpaid, or Partial."""
    cfg = ctx.obj["config"]
    known = all_firm_names(cfg)
    if firm not in known:
        raise click.ClickException(f"Firm '{firm}' not found. Available: {known}")

    if payment_date:
        from datetime import date as _date
        try:
            _date.fromisoformat(payment_date)
        except ValueError:
            raise click.ClickException(f"Invalid date: {payment_date}. Use YYYY-MM-DD.")

    try:
        updated = mark_payment(firm, invoice_number, status, payment_date, notes)
    except (ValueError, FileNotFoundError) as e:
        raise click.ClickException(str(e))

    click.echo(f"Payment updated: {firm}")
    click.echo(f"  Invoice:      {invoice_number}")
    click.echo(f"  Case:         {updated.get('case_caption', '')}")
    click.echo(f"  Amount:       ${float(updated.get('charge_amount') or 0):,.2f}")
    click.echo(f"  Status:       {updated.get('paid_status', '')}")
    click.echo(f"  Payment date: {updated.get('payment_date', '')}")
    if notes:
        click.echo(f"  Notes:        {notes}")


if __name__ == "__main__":
    cli()

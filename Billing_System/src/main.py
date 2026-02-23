"""CLI entrypoint for the law firm billing system."""

import click

from src.config import load_config
from src.dataset import COLUMNS, VALID_CASE_STATUSES
from src.payment import VALID_STATUSES
from src.services import case_service, doc_service, payment_service


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
    result = case_service.init_datasets(firm=firm, force=force, config=ctx.obj["config"])
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


@cli.command("validate-dataset")
@click.option("--firm", default=None, help="Firm name (omit to validate all firms).")
@click.pass_context
def validate_dataset_cmd(ctx, firm):
    """Validate master_cases.xlsx for each firm."""
    result = case_service.validate_datasets(firm=firm, config=ctx.obj["config"])
    click.echo(result.message)
    if not result.success:
        raise SystemExit(1)


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
    result = case_service.add_or_update_case(
        firm, appearance_date, index_number, case_caption, charge_amount,
        court=court, outcome=outcome, case_status=case_status, notes=notes,
        config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 4: invoice numbering ────────────────────────────────────────


@cli.command("assign-invoices")
@click.option("--firm", default=None, help="Firm name (omit to assign for all firms).")
@click.pass_context
def assign_invoices(ctx, firm):
    """Assign invoice numbers to cases that don't have one yet."""
    result = case_service.assign_invoices(firm=firm, config=ctx.obj["config"])
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 5: generate daily invoice ───────────────────────────────────


@cli.command("generate-daily")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--index", "index_number", required=True, help="Case index number.")
@click.option("--date", "appearance_date", required=True, help="Appearance date (YYYY-MM-DD).")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def generate_daily(ctx, firm, index_number, appearance_date, keep_docx):
    """Generate a per-diem invoice PDF for a specific case."""
    result = doc_service.generate_daily(
        firm, index_number, appearance_date,
        keep_docx=keep_docx, config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 7: weekly statement ─────────────────────────────────────────


@cli.command("generate-weekly")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--week-of", required=True, help="Any date within the week (YYYY-MM-DD). Mon-Fri range is computed.")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def generate_weekly(ctx, firm, week_of, keep_docx):
    """Generate a weekly statement of account for a firm."""
    result = doc_service.generate_weekly(
        firm, week_of, keep_docx=keep_docx, config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 8: monthly statement ───────────────────────────────────────


@cli.command("generate-monthly")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--year", required=True, type=int, help="Year (e.g. 2026).")
@click.option("--month", required=True, type=int, help="Month number (1-12).")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def generate_monthly(ctx, firm, year, month, keep_docx):
    """Generate a monthly statement of account for a firm."""
    result = doc_service.generate_monthly(
        firm, year, month, keep_docx=keep_docx, config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 9: master ledger export ────────────────────────────────────


@cli.command("export-ledger")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--asof", default=None, help="As-of date (YYYY-MM-DD). Defaults to today.")
@click.option("--no-xlsx", is_flag=True, help="Skip XLSX generation (PDF only).")
@click.option("--keep-docx", is_flag=True, help="Keep intermediate .docx file.")
@click.pass_context
def export_ledger_cmd(ctx, firm, asof, no_xlsx, keep_docx):
    """Export a firm's full-history master ledger (PDF + XLSX)."""
    result = doc_service.export_ledger(
        firm, as_of=asof, xlsx=not no_xlsx,
        keep_docx=keep_docx, config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 10: legacy import ──────────────────────────────────────────


@cli.command("import-legacy")
@click.option("--firm", required=True, help="Law firm name (must match config).")
@click.option("--file", "file_path", required=True, type=click.Path(exists=True), help="Path to legacy monthly invoice .docx file.")
@click.pass_context
def import_legacy(ctx, firm, file_path):
    """Import cases from a legacy monthly invoice .docx into a firm's dataset."""
    result = case_service.import_legacy(firm, file_path, config=ctx.obj["config"])
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


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
    result = payment_service.mark_paid(
        firm, invoice_number, status,
        payment_date=payment_date, notes=notes,
        config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


# ── Phase 14: edit case field ─────────────────────────────────────────


EDITABLE_FIELDS = sorted(["charge_amount", "court", "outcome", "case_status", "notes"])


@cli.command("edit-case")
@click.option("--firm", required=True, help="Law firm name.")
@click.option("--index", "index_number", required=True, help="Case index number.")
@click.option("--date", "appearance_date", required=True, help="Appearance date (YYYY-MM-DD).")
@click.option("--field", "field_name", required=True,
              type=click.Choice(EDITABLE_FIELDS, case_sensitive=False),
              help="Field to edit.")
@click.option("--value", "new_value", required=True, help="New value for the field.")
@click.option("--reason", default=None, help="Reason for edit (required for charge_amount after invoice sent).")
@click.pass_context
def edit_case(ctx, firm, index_number, appearance_date, field_name, new_value, reason):
    """Edit a single field on an existing case (with audit logging)."""
    result = case_service.edit_case_field(
        firm, index_number, appearance_date, field_name, new_value,
        reason=reason, config=ctx.obj["config"],
    )
    if not result.success:
        raise click.ClickException(result.message)
    click.echo(result.message)


if __name__ == "__main__":
    cli()

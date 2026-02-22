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


# ── Phase 7 placeholder ──────────────────────────────────────────────


@cli.command("generate-weekly")
def generate_weekly():
    """(Phase 7) Generate a weekly invoice batch."""
    click.echo("Not yet implemented — coming in Phase 7.")


# ── Phase 8 placeholder ──────────────────────────────────────────────


@cli.command("generate-monthly")
def generate_monthly():
    """(Phase 8) Generate a monthly invoice summary."""
    click.echo("Not yet implemented — coming in Phase 8.")


# ── Phase 9 placeholder ──────────────────────────────────────────────


@cli.command("export-ledger")
def export_ledger():
    """(Phase 9) Export the billing ledger to Excel."""
    click.echo("Not yet implemented — coming in Phase 9.")


# ── Phase 11 placeholder ─────────────────────────────────────────────


@cli.command("mark-paid")
def mark_paid():
    """(Phase 11) Mark an invoice as paid."""
    click.echo("Not yet implemented — coming in Phase 11.")


if __name__ == "__main__":
    cli()

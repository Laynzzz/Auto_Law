"""Case-data service — init, validate, add/update, assign invoices, import legacy, extract firms, bulk import."""

from __future__ import annotations

import json
from pathlib import Path

from src.config import CONFIG_PATH, load_config
from src.audit_log import append_audit
from src.dataset import (
    COLUMNS,
    VALID_CASE_STATUSES,
    all_firm_names,
    create_workbook,
    dataset_path,
    find_row_by_key,
    load_dataset,
    upsert_row,
    validate_dataset,
)
from src.invoice_number import assign_invoice_numbers
from src.legacy_import import import_legacy_invoice
from src.services import ServiceResult


# ── Helpers ──────────────────────────────────────────────────────────


def _resolve_config(config: dict | None) -> dict:
    if config is None:
        return load_config()
    return config


def _validate_firm(firm: str, config: dict) -> str | None:
    """Return an error message if *firm* is not in config, else None."""
    known = all_firm_names(config)
    if firm not in known:
        return f"Firm '{firm}' not found. Available: {known}"
    return None


# ── Public API ───────────────────────────────────────────────────────


def init_datasets(
    firm: str | None = None,
    force: bool = False,
    config: dict | None = None,
) -> ServiceResult:
    """Create master_cases.xlsx for one firm or all firms.

    Returns created file paths in ``data["created"]``.
    """
    config = _resolve_config(config)
    firms = [firm] if firm else all_firm_names(config)

    created: list[str] = []
    for name in firms:
        try:
            path = create_workbook(name, overwrite=force)
            created.append(str(path))
        except FileExistsError as exc:
            return ServiceResult(success=False, message=str(exc))

    lines = [f"Created: {p}" for p in created]
    lines.append(f"\n  Sheet: cases")
    lines.append(f"  Columns ({len(COLUMNS)}): {', '.join(COLUMNS)}")
    lines.append(f"  Initialized {len(firms)} firm(s).")

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={"created": created},
    )


def validate_datasets(
    firm: str | None = None,
    config: dict | None = None,
) -> ServiceResult:
    """Validate master_cases.xlsx for one firm or all firms.

    Returns per-firm errors and row counts in ``data["firms"]``.
    """
    config = _resolve_config(config)
    firms = [firm] if firm else all_firm_names(config)

    results: dict[str, dict] = {}
    total_errors = 0
    lines: list[str] = []

    for name in firms:
        path = dataset_path(name)
        lines.append(f"--- {name} ---")
        lines.append(f"  File: {path}")

        errors = validate_dataset(name)
        if errors:
            lines.append(f"  FAILED - {len(errors)} error(s):")
            for err in errors:
                lines.append(f"    - {err}")
            total_errors += len(errors)
            results[name] = {"errors": errors, "row_count": None}
        else:
            rows = load_dataset(name)
            lines.append(f"  OK - {len(rows)} data row(s)")
            results[name] = {"errors": [], "row_count": len(rows)}
        lines.append("")

    if total_errors:
        lines.append(f"Total errors across all firms: {total_errors}")
        return ServiceResult(
            success=False,
            message="\n".join(lines),
            data={"firms": results},
        )

    lines.append(f"All {len(firms)} firm(s) validated OK.")
    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={"firms": results},
    )


def add_or_update_case(
    firm: str,
    appearance_date: str,
    index_number: str,
    case_caption: str,
    charge_amount: float,
    court: str | None = None,
    outcome: str | None = None,
    case_status: str | None = None,
    notes: str | None = None,
    config: dict | None = None,
) -> ServiceResult:
    """Add or update a case in a firm's dataset.

    Returns the action taken and the row data in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

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
    except FileNotFoundError as exc:
        return ServiceResult(success=False, message=str(exc))

    lines = [f"Case {action}: {firm}"]
    lines.append(f"  index_number:    {index_number}")
    lines.append(f"  appearance_date: {appearance_date}")
    lines.append(f"  case_caption:    {case_caption}")
    lines.append(f"  charge_amount:   {charge_amount}")
    if court:
        lines.append(f"  court:           {court}")
    if outcome:
        lines.append(f"  outcome:         {outcome}")
    if case_status:
        lines.append(f"  case_status:     {case_status}")
    if notes:
        lines.append(f"  notes:           {notes}")

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={"action": action, "firm": firm, "row": row_data},
    )


def assign_invoices(
    firm: str | None = None,
    config: dict | None = None,
) -> ServiceResult:
    """Assign invoice numbers to cases that don't have one yet.

    Returns assigned invoice numbers per firm in ``data["assigned"]``.
    """
    config = _resolve_config(config)
    firms = [firm] if firm else all_firm_names(config)

    all_assigned: dict[str, list[str]] = {}
    lines: list[str] = []

    for name in firms:
        lines.append(f"--- {name} ---")
        try:
            assigned = assign_invoice_numbers(name, config)
        except FileNotFoundError as exc:
            return ServiceResult(success=False, message=str(exc))

        all_assigned[name] = assigned
        if assigned:
            for inv in assigned:
                lines.append(f"  Assigned: {inv}")
            lines.append(f"  Total new: {len(assigned)}")
        else:
            lines.append("  No cases need invoice numbers.")
        lines.append("")

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={"assigned": all_assigned},
    )


def import_legacy(
    firm: str,
    file_path: str,
    config: dict | None = None,
) -> ServiceResult:
    """Import cases from a legacy monthly invoice .docx into a firm's dataset.

    Returns import results in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    try:
        results = import_legacy_invoice(firm, file_path)
    except (FileNotFoundError, ValueError) as exc:
        return ServiceResult(success=False, message=str(exc))

    if not results:
        return ServiceResult(
            success=True,
            message="No cases found in the file.",
            data={"results": [], "inserted": 0, "updated": 0},
        )

    inserted = 0
    updated = 0
    lines: list[str] = []

    for action, case in results:
        label = "NEW" if action == "inserted" else "UPD"
        lines.append(
            f"  [{label}] {case['appearance_date']} | "
            f"{case['index_number']} | {case['case_caption']} | "
            f"${case['charge_amount']:.2f}"
        )
        if action == "inserted":
            inserted += 1
        else:
            updated += 1

    lines.append(f"\nImported {len(results)} case(s): {inserted} new, {updated} updated.")

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={
            "results": [(a, c) for a, c in results],
            "inserted": inserted,
            "updated": updated,
        },
    )


# ── Phase 14: edit case field ────────────────────────────────────────

EDITABLE_FIELDS = {
    "charge_amount": "EDIT_CHARGE",
    "court":         "EDIT_COURT",
    "outcome":       "EDIT_OUTCOME",
    "case_status":   "EDIT_STATUS",
    "notes":         "EDIT_NOTES",
}


def edit_case_field(
    firm: str,
    index_number: str,
    appearance_date: str,
    field_name: str,
    new_value: str,
    reason: str | None = None,
    config: dict | None = None,
) -> ServiceResult:
    """Edit a single field on an existing case, with audit logging.

    Returns old and new values in ``data``.
    """
    config = _resolve_config(config)

    err = _validate_firm(firm, config)
    if err:
        return ServiceResult(success=False, message=err)

    if field_name not in EDITABLE_FIELDS:
        allowed = ", ".join(sorted(EDITABLE_FIELDS))
        return ServiceResult(
            success=False,
            message=f"Field '{field_name}' is not editable. Allowed: {allowed}",
        )

    row = find_row_by_key(firm, index_number, appearance_date)
    if row is None:
        return ServiceResult(
            success=False,
            message=(
                f"Case not found: index={index_number}, "
                f"date={appearance_date} in firm '{firm}'."
            ),
        )

    # Require reason for charge_amount edits after invoice has been sent
    invoice_sent = row.get("invoice_sent_date")
    if field_name == "charge_amount" and invoice_sent is not None and str(invoice_sent).strip():
        if not reason:
            return ServiceResult(
                success=False,
                message=(
                    "Reason is required when editing charge_amount "
                    "after the invoice has been sent. Use --reason."
                ),
            )

    old_value = row.get(field_name)

    # Type coercion
    coerced_value: object = new_value
    if field_name == "charge_amount":
        try:
            coerced_value = float(new_value)
        except ValueError:
            return ServiceResult(
                success=False,
                message=f"Invalid charge_amount: '{new_value}' is not a number.",
            )
    elif field_name == "case_status":
        if new_value not in VALID_CASE_STATUSES:
            return ServiceResult(
                success=False,
                message=(
                    f"Invalid case_status '{new_value}'. "
                    f"Must be one of: {sorted(VALID_CASE_STATUSES)}"
                ),
            )

    # Write update
    upsert_row(firm, {
        "index_number": index_number,
        "appearance_date": appearance_date,
        field_name: coerced_value,
    })

    # Write audit log
    action = EDITABLE_FIELDS[field_name]
    append_audit(
        firm=firm,
        index_number=index_number,
        appearance_date=appearance_date,
        action=action,
        field_name=field_name,
        old_value=old_value,
        new_value=coerced_value,
        reason=reason,
    )

    return ServiceResult(
        success=True,
        message=(
            f"Updated {field_name} on case {index_number} ({appearance_date}):\n"
            f"  {old_value} -> {coerced_value}"
            + (f"\n  Reason: {reason}" if reason else "")
        ),
        data={
            "action": action,
            "old_value": old_value,
            "new_value": coerced_value,
        },
    )


# ── Phase 18: extract firm info from invoices ────────────────────────


def extract_firms(
    invoices_dir: str,
    output_path: str | None = None,
    config: dict | None = None,
) -> ServiceResult:
    """Scan invoice folders, extract firm metadata, write preview JSON.

    Args:
        invoices_dir: Path to the invoices root (e.g. "invoice/2026 Invoices").
        output_path: Where to write the preview JSON. Defaults to data/extracted_firms.json.
        config: Optional config dict.

    Returns:
        ServiceResult with count of firms extracted and any warnings.
    """
    from src.firm_extractor import scan_all_firms

    config = _resolve_config(config)

    inv_path = Path(invoices_dir)
    if not inv_path.is_dir():
        return ServiceResult(
            success=False,
            message=f"Directory not found: {invoices_dir}",
        )

    firms, warnings = scan_all_firms(inv_path)

    # Determine output path
    if output_path is None:
        data_dir = Path(config.get("paths", {}).get("data_dir", "data"))
        data_dir.mkdir(parents=True, exist_ok=True)
        out = data_dir / "extracted_firms.json"
    else:
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)

    # Write preview JSON
    with open(out, "w", encoding="utf-8") as f:
        json.dump(firms, f, indent=2, ensure_ascii=False)

    # Build summary
    lines: list[str] = []
    lines.append(f"Extracted {len(firms)} firm(s) from {invoices_dir}")
    if warnings:
        lines.append(f"\nWarnings ({len(warnings)}):")
        for w in warnings:
            lines.append(f"  - {w}")
    lines.append(f"\nPreview written to: {out}")

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={
            "firms_count": len(firms),
            "warnings": warnings,
            "output_path": str(out),
        },
    )


# ── Phase 19: bulk import ────────────────────────────────────────────


def bulk_import(
    firms_json: str,
    invoices_dir: str,
    config: dict | None = None,
) -> ServiceResult:
    """Add extracted firms to config and create blank datasets.

    Args:
        firms_json: Path to the extracted firms JSON (e.g. data/extracted_firms.json).
        invoices_dir: Path to the invoices root (used only for validation).
        config: Optional config dict.

    Returns:
        ServiceResult with summary of firms added and datasets created.
    """
    config = _resolve_config(config)

    # 1. Load extracted firms
    firms_path = Path(firms_json)
    if not firms_path.is_file():
        return ServiceResult(success=False, message=f"File not found: {firms_json}")

    with open(firms_path, "r", encoding="utf-8") as f:
        extracted = json.load(f)

    inv_dir = Path(invoices_dir)
    if not inv_dir.is_dir():
        return ServiceResult(success=False, message=f"Directory not found: {invoices_dir}")

    # 2. Build set of existing firms (case-insensitive) and initials
    existing_names = {f["name"].lower() for f in config["firms"]}
    existing_initials = {f["initials"].upper() for f in config["firms"]}

    # 3. Deduplicate new firms by folder_name (skip those already in config)
    new_firms: list[dict] = []
    seen_folders: set[str] = set()

    for entry in extracted:
        folder = entry.get("folder_name", "")
        if not folder:
            continue
        if folder.lower() in existing_names:
            continue
        if folder.lower() in seen_folders:
            continue
        seen_folders.add(folder.lower())
        new_firms.append(entry)

    # 4. Resolve initials conflicts
    used_initials = set(existing_initials)

    for firm in new_firms:
        initials = firm.get("initials", "").upper()
        if not initials:
            # Generate from folder name words
            words = firm["folder_name"].split()
            initials = "".join(w[0] for w in words if w).upper()[:3]

        if initials in used_initials:
            # Append letters from folder name until unique
            alpha_chars = [c.upper() for c in firm["folder_name"] if c.isalpha()]
            resolved = False
            for ch in alpha_chars:
                candidate = initials + ch
                if candidate not in used_initials:
                    initials = candidate
                    resolved = True
                    break
            if not resolved:
                for i in range(1, 100):
                    candidate = initials + str(i)
                    if candidate not in used_initials:
                        initials = candidate
                        break

        firm["_resolved_initials"] = initials
        used_initials.add(initials)

    # 5. Add new firms to config and save
    for firm in new_firms:
        config["firms"].append({
            "name": firm["folder_name"],
            "initials": firm["_resolved_initials"],
            "contact_name": firm.get("contact_name", ""),
            "address_1": firm.get("address_1", ""),
            "address_2": firm.get("address_2", ""),
            "phone": firm.get("phone", ""),
            "billing_email": firm.get("billing_email", ""),
            "cc_emails": firm.get("cc_emails", []),
        })

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    # 6. Create blank datasets for new firms
    datasets_created = 0
    for firm in new_firms:
        try:
            create_workbook(firm["folder_name"])
            datasets_created += 1
        except FileExistsError:
            pass  # already exists

    # 7. Build summary
    lines = [
        "Bulk import complete.",
        f"  Firms added to config: {len(new_firms)}",
        f"  Total config firms: {len(config['firms'])}",
        f"  Datasets created: {datasets_created}",
    ]

    return ServiceResult(
        success=True,
        message="\n".join(lines),
        data={
            "firms_added": len(new_firms),
            "datasets_created": datasets_created,
        },
    )

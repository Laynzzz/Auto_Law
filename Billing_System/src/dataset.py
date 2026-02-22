"""Master dataset (source-of-truth) schema and operations.

Each law firm has its own dataset at  invoice/{FirmName}/master_{FirmName}.xlsx
(sheet: 'cases').  Each row represents one coverage appearance / billable event.

Unique key within a firm file:  (index_number, appearance_date)
"""

from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.config import CONFIG_PATH, load_config

# ── Schema ────────────────────────────────────────────────────────────

COLUMNS = [
    "appearance_date",
    "invoice_number",
    "index_number",
    "case_caption",
    "court",
    "outcome",
    "case_status",
    "charge_amount",
    "invoice_sent_date",
    "paid_status",
    "payment_date",
    "notes",
]

REQUIRED_COLUMNS = [
    "case_caption",
    "index_number",
    "appearance_date",
    "charge_amount",
]

VALID_CASE_STATUSES = {"Open", "Adjourned", "Closed", "Settled", "Dismissed"}

VALID_PAID_STATUSES = {"Paid", "Unpaid", "Partial"}

# Unique key within a per-firm file
UNIQUE_KEY_COLS = ("index_number", "appearance_date")

# ── Paths ─────────────────────────────────────────────────────────────

PROJECT_ROOT = CONFIG_PATH.parent.parent


def dataset_path(firm_name: str) -> Path:
    """Return path to a firm's master dataset: invoice/{firm_name}/master_{firm_name}.xlsx"""
    return PROJECT_ROOT / "invoice" / firm_name / f"master_{firm_name}.xlsx"


def all_firm_names(config: dict | None = None) -> list[str]:
    """Return list of firm names from config."""
    if config is None:
        config = load_config()
    return [f["name"] for f in config["firms"]]


# ── Create ────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)


def create_workbook(firm_name: str, overwrite: bool = False) -> Path:
    """Create a new master_cases.xlsx for a firm with the 'cases' sheet and headers."""
    path = dataset_path(firm_name)

    if path.exists() and not overwrite:
        raise FileExistsError(
            f"Dataset already exists: {path}\n"
            "Use --force to overwrite (this will erase all data)."
        )

    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "cases"

    # Write header row
    for col_idx, name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    # Auto-width based on header length
    for col_idx, name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(name) + 4, 14)

    # Freeze the header row
    ws.freeze_panes = "A2"

    wb.save(path)
    return path


def create_all_workbooks(config: dict | None = None, overwrite: bool = False) -> list[Path]:
    """Create master_cases.xlsx for every firm in config. Returns list of created paths."""
    created: list[Path] = []
    for name in all_firm_names(config):
        path = create_workbook(name, overwrite=overwrite)
        created.append(path)
    return created


# ── Load ──────────────────────────────────────────────────────────────


def load_dataset(firm_name: str) -> list[dict]:
    """Load all rows from a firm's 'cases' sheet as a list of dicts."""
    path = dataset_path(firm_name)

    if not path.exists():
        raise FileNotFoundError(
            f"Dataset not found: {path}\n"
            f"Run 'python -m src.main init-dataset' first."
        )

    wb = load_workbook(path)
    ws = wb["cases"]

    headers = [cell.value for cell in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        rows.append(dict(zip(headers, row)))

    wb.close()
    return rows


# ── Validate ──────────────────────────────────────────────────────────


def validate_dataset(firm_name: str) -> list[str]:
    """Validate a firm's dataset file. Returns list of error messages (empty = OK)."""
    path = dataset_path(firm_name)
    errors: list[str] = []

    if not path.exists():
        return [f"Dataset file not found: {path}"]

    wb = load_workbook(path)

    # Check sheet exists
    if "cases" not in wb.sheetnames:
        wb.close()
        return ["Missing required sheet 'cases'"]

    ws = wb["cases"]
    headers = [cell.value for cell in ws[1]]

    # Check all expected columns present
    missing_cols = [c for c in COLUMNS if c not in headers]
    if missing_cols:
        errors.append(f"Missing columns: {missing_cols}")
        wb.close()
        return errors  # Can't validate rows without correct headers

    # Validate each data row
    seen_keys: set[tuple] = set()
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        row_dict = dict(zip(headers, row))

        if all(v is None for v in row):
            continue

        # Required fields
        for col in REQUIRED_COLUMNS:
            if row_dict.get(col) is None or str(row_dict[col]).strip() == "":
                errors.append(f"Row {row_num}: missing required field '{col}'")

        # appearance_date format
        ad = row_dict.get("appearance_date")
        if ad is not None:
            if isinstance(ad, datetime):
                pass  # openpyxl parses dates as datetime
            elif isinstance(ad, date):
                pass
            elif isinstance(ad, str):
                try:
                    datetime.strptime(ad, "%Y-%m-%d")
                except ValueError:
                    errors.append(
                        f"Row {row_num}: appearance_date '{ad}' is not YYYY-MM-DD"
                    )

        # charge_amount must be numeric
        amt = row_dict.get("charge_amount")
        if amt is not None and not isinstance(amt, (int, float)):
            try:
                float(amt)
            except (ValueError, TypeError):
                errors.append(
                    f"Row {row_num}: charge_amount '{amt}' is not a number"
                )

        # case_status validation
        cs = row_dict.get("case_status")
        if cs is not None and str(cs).strip() != "":
            if str(cs).strip() not in VALID_CASE_STATUSES:
                errors.append(
                    f"Row {row_num}: case_status '{cs}' not in {VALID_CASE_STATUSES}"
                )

        # paid_status validation
        ps = row_dict.get("paid_status")
        if ps is not None and str(ps).strip() != "":
            if str(ps).strip() not in VALID_PAID_STATUSES:
                errors.append(
                    f"Row {row_num}: paid_status '{ps}' not in {VALID_PAID_STATUSES}"
                )

        # Unique key check (within a per-firm file: index_number + appearance_date)
        key = (
            str(row_dict.get("index_number", "")).strip().lower(),
            str(row_dict.get("appearance_date", "")),
        )
        if key in seen_keys:
            errors.append(f"Row {row_num}: duplicate key {key}")
        seen_keys.add(key)

    wb.close()
    return errors


# ── Lookup ────────────────────────────────────────────────────────────


def find_row_by_key(
    firm_name: str,
    index_number: str,
    appearance_date: str | date,
    rows: list[dict] | None = None,
) -> dict | None:
    """Find a row by its unique key within a firm's dataset. Returns the dict or None."""
    if rows is None:
        rows = load_dataset(firm_name)

    target_idx = index_number.strip().lower()
    target_date = str(appearance_date)

    for row in rows:
        if (
            str(row.get("index_number", "")).strip().lower() == target_idx
            and str(row.get("appearance_date", "")) == target_date
        ):
            return row

    return None


# ── Upsert ────────────────────────────────────────────────────────────


def _match_key(ws, headers: list[str], index_number: str, appearance_date: str) -> int | None:
    """Return the Excel row number of a matching key, or None."""
    idx_col = headers.index("index_number")
    date_col = headers.index("appearance_date")
    target_idx = index_number.strip().lower()
    target_date = str(appearance_date)

    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        row_idx = str(row[idx_col] or "").strip().lower()
        row_date = str(row[date_col] or "")
        if row_idx == target_idx and row_date == target_date:
            return row_num
    return None


# ── Query ─────────────────────────────────────────────────────────────


def _to_date(val) -> date | None:
    """Coerce a value to a date object, or None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.strptime(str(val).split(" ")[0], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def query_by_date_range(
    firm_name: str,
    start: date,
    end: date,
    rows: list[dict] | None = None,
) -> list[dict]:
    """Return rows whose appearance_date falls within [start, end] inclusive."""
    if rows is None:
        rows = load_dataset(firm_name)

    result = []
    for row in rows:
        d = _to_date(row.get("appearance_date"))
        if d is not None and start <= d <= end:
            result.append(row)

    # Sort by appearance_date
    result.sort(key=lambda r: _to_date(r.get("appearance_date")) or date.min)
    return result


def week_range(ref_date: date) -> tuple[date, date]:
    """Return (monday, friday) of the business week containing ref_date."""
    monday = ref_date - timedelta(days=ref_date.weekday())  # weekday: Mon=0
    friday = monday + timedelta(days=4)
    return monday, friday


def upsert_row(firm_name: str, row_data: dict) -> str:
    """Insert or update a row in a firm's dataset.

    row_data keys should match COLUMNS (extras are ignored, missing become None).
    Uses (index_number, appearance_date) as the unique key.

    Returns "inserted" or "updated".
    """
    path = dataset_path(firm_name)

    if not path.exists():
        raise FileNotFoundError(
            f"Dataset not found: {path}\n"
            "Run 'python -m src.main init-dataset' first."
        )

    wb = load_workbook(path)
    ws = wb["cases"]
    headers = [cell.value for cell in ws[1]]

    idx_num = str(row_data.get("index_number", ""))
    app_date = str(row_data.get("appearance_date", ""))
    existing_row = _match_key(ws, headers, idx_num, app_date)

    if existing_row is not None:
        # Update existing row — overwrite only fields that are provided
        for col_name, value in row_data.items():
            if col_name in headers:
                col_idx = headers.index(col_name) + 1  # 1-based
                ws.cell(row=existing_row, column=col_idx, value=value)
        wb.save(path)
        wb.close()
        return "updated"
    else:
        # Append new row
        new_row = [row_data.get(col) for col in headers]
        ws.append(new_row)
        wb.save(path)
        wb.close()
        return "inserted"

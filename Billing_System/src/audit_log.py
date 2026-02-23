"""Centralized audit log for field-level edits on cases.

Writes to {data_root}/invoice/audit_log.csv — shared across all firms.
Each row records who changed what, when, and why.
"""

import csv
import os
import socket
from datetime import datetime
from pathlib import Path

from src.dataset import get_data_root

AUDIT_COLUMNS = [
    "timestamp",
    "user",
    "hostname",
    "action",
    "firm",
    "case_key",
    "field_name",
    "old_value",
    "new_value",
    "reason",
]


def _audit_log_path() -> Path:
    return get_data_root() / "invoice" / "audit_log.csv"


def append_audit(
    firm: str,
    index_number: str,
    appearance_date: str,
    action: str,
    field_name: str,
    old_value,
    new_value,
    reason: str | None = None,
) -> None:
    """Append one audit row to the shared audit log CSV.

    Creates the file with headers if it doesn't exist yet.
    """
    log_path = _audit_log_path()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    file_exists = log_path.exists()

    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(AUDIT_COLUMNS)
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            os.getlogin(),
            socket.gethostname(),
            action,
            firm,
            f"{index_number}|{appearance_date}",
            field_name,
            str(old_value) if old_value is not None else "",
            str(new_value) if new_value is not None else "",
            reason or "",
        ])

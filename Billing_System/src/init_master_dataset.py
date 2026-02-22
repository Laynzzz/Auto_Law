"""Standalone script to create master_cases.xlsx per firm.

Usage:
    python -m src.init_master_dataset [--firm "ABC Law"] [--force]
"""

import argparse

from src.config import load_config
from src.dataset import COLUMNS, all_firm_names, create_workbook


def main():
    parser = argparse.ArgumentParser(
        description="Create master_cases.xlsx per firm with the 'cases' sheet and headers."
    )
    parser.add_argument(
        "--firm",
        default=None,
        help="Firm name (omit to init all firms).",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing file (erases all data).",
    )
    args = parser.parse_args()

    cfg = load_config()
    firms = [args.firm] if args.firm else all_firm_names(cfg)

    for name in firms:
        try:
            path = create_workbook(name, overwrite=args.force)
            print(f"Created: {path}")
        except FileExistsError as e:
            print(f"ERROR: {e}")
            raise SystemExit(1)

    print(f"\n  Sheet: cases")
    print(f"  Columns ({len(COLUMNS)}): {', '.join(COLUMNS)}")
    print(f"  Initialized {len(firms)} firm(s).")


if __name__ == "__main__":
    main()

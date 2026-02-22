"""Standalone script to validate master_cases.xlsx per firm.

Usage:
    python -m src.validate_dataset [--firm "ABC Law"]
"""

import argparse

from src.config import load_config
from src.dataset import all_firm_names, dataset_path, load_dataset, validate_dataset


def main():
    parser = argparse.ArgumentParser(
        description="Validate master_cases.xlsx for each firm."
    )
    parser.add_argument(
        "--firm",
        default=None,
        help="Firm name (omit to validate all firms).",
    )
    args = parser.parse_args()

    cfg = load_config()
    firms = [args.firm] if args.firm else all_firm_names(cfg)
    total_errors = 0

    for name in firms:
        path = dataset_path(name)
        print(f"--- {name} ---")
        print(f"  File: {path}")

        errors = validate_dataset(name)
        if errors:
            print(f"  FAILED - {len(errors)} error(s):")
            for err in errors:
                print(f"    - {err}")
            total_errors += len(errors)
        else:
            rows = load_dataset(name)
            print(f"  OK - {len(rows)} data row(s)")
        print()

    if total_errors:
        print(f"Total errors across all firms: {total_errors}")
        raise SystemExit(1)
    else:
        print(f"All {len(firms)} firm(s) validated OK.")


if __name__ == "__main__":
    main()

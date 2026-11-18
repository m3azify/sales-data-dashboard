# -*- coding: utf-8 -*-
"""Check that an input sales CSV exists and has a usable header.

Why this is script 01 (and stdlib-only): it must run BEFORE `pip install`, so a
reviewer can be told exactly what to download without needing any package. It
never downloads anything itself (no network, no API keys).

Exit codes: 0 = input ready, 1 = missing/invalid (instructions printed).
Run from project root: python src/01_check_data.py
"""
from __future__ import annotations

import csv
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = PROJECT_ROOT / "data" / "raw"
REAL_PATH = RAW_DIR / "superstore.csv"
SYNTHETIC_PATH = RAW_DIR / "synthetic_superstore.csv"

# Columns we refuse to work without. Optional ones (Country, Postal Code) are kept
# if present but are not required.
REQUIRED_COLUMNS = [
    "Row ID", "Order ID", "Order Date", "Ship Date", "Ship Mode",
    "Customer ID", "Customer Name", "Segment", "City", "State", "Region",
    "Product ID", "Category", "Sub-Category", "Product Name",
    "Sales", "Quantity", "Discount", "Profit",
]


def _normalise_header(name: str) -> str:
    return " ".join(name.strip().lower().split())


def _read_header_and_count(path: Path) -> tuple[list[str], int]:
    """Return (header, data_row_count). Raises ValueError on a bad file."""
    # utf-8-sig strips a BOM if present (common in Kaggle exports).
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with path.open("r", encoding=enc, newline="") as f:
                reader = csv.reader(f)
                try:
                    header = next(reader)
                except StopIteration:
                    raise ValueError("the file is empty")
                count = 0
                bad_rows = []
                for row in reader:
                    if not row:  # skip a trailing blank line
                        continue
                    count += 1
                    if len(row) != len(header) and len(bad_rows) < 3:
                        bad_rows.append((count, len(row)))
                if bad_rows:
                    raise ValueError(
                        "rows have a different number of columns than the header "
                        "(first offenders: " + ", ".join(f"row {r} has {c} fields" for r, c in bad_rows) + ")"
                    )
                return [h.strip() for h in header], count
        except UnicodeDecodeError:
            continue
    raise ValueError("could not decode the file as utf-8 or cp1252 (is it a zip? unzip it first)")


def _print_instructions() -> None:
    print("-" * 72)
    print("INPUT DATA NOT FOUND - download instructions")
    print("-" * 72)
    print(f"Expected folder : {RAW_DIR}")
    print("Expected file   : superstore.csv  (the Kaggle 'Sample - Superstore.csv')")
    print()
    print("Official source (free Kaggle account required):")
    print("  1. Sign in at https://www.kaggle.com")
    print("  2. Open https://www.kaggle.com/datasets/vivek468/superstore-dataset-final")
    print("  3. Click Download, unzip, rename 'Sample - Superstore.csv' -> 'superstore.csv'")
    print("  4. Place it in the folder above.")
    print()
    print("No Kaggle account? Generate a synthetic demo instead (clearly labelled):")
    print("  python src/00_generate_sample.py")
    print("  (writes data/raw/synthetic_superstore.csv; the pipeline prefers the real")
    print("   file if both are present)")
    print("-" * 72)


def _validate(path: Path, label: str) -> int:
    header, count = _read_header_and_count(path)
    norm = {_normalise_header(h): h for h in header}
    missing = [c for c in REQUIRED_COLUMNS if _normalise_header(c) not in norm]
    if missing:
        print(f"[01] {label}: {path}")
        print(f"     FAIL - missing required columns: {missing}")
        print("     See the download instructions below.")
        return 1
    size_mb = path.stat().st_size / (1024 * 1024)
    print(f"[01] {label}: {path}")
    print(f"     OK - {count:,} data rows, {len(header)} columns, {size_mb:.2f} MB")
    print(f"     Columns: {', '.join(header)}")
    return 0


def main() -> int:
    print("[01] Checking input data...")
    print(f"     Project root: {PROJECT_ROOT}")

    if REAL_PATH.exists() and SYNTHETIC_PATH.exists():
        print("     Both files present; the real dataset takes priority.")
    elif REAL_PATH.exists():
        pass
    elif SYNTHETIC_PATH.exists():
        print("     NOTE: only the SYNTHETIC dataset is present - outputs will be labelled synthetic.")
    else:
        _print_instructions()
        return 1

    # Real file wins if it exists, but never fall back silently if it is invalid.
    if REAL_PATH.exists():
        try:
            rc = _validate(REAL_PATH, "REAL dataset (superstore.csv)")
        except ValueError as exc:
            print(f"[01] Real file found but INVALID: {exc}")
            print("     Not falling back to synthetic - fix or remove the real file first.")
            _print_instructions()
            return 1
        if rc != 0:
            _print_instructions()
            return 1
        return 0

    # Synthetic path only.
    try:
        rc = _validate(SYNTHETIC_PATH, "SYNTHETIC dataset (synthetic_superstore.csv)")
    except ValueError as exc:
        print(f"[01] Synthetic file found but INVALID: {exc}")
        _print_instructions()
        return 1
    if rc != 0:
        _print_instructions()
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

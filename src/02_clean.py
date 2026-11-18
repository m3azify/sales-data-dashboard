# -*- coding: utf-8 -*-
"""Clean the raw sales CSV into analysis-ready data.

Why a separate script: cleaning is the step a business owner cannot see but
everything downstream depends on. Keeping it isolated means we can rerun it
without touching reports, and a reviewer can audit one file.

Rules (each is printed with a count so nothing is silent):
  - Drop exact duplicates (after trimming whitespace, so ' Consumer ' == 'Consumer').
  - Parse dates from explicit formats only (no guessing day/month order).
  - Drop rows missing critical fields (Order ID, Order/Ship Date, Sales, Profit).
  - Fill optional nulls (name -> 'Unknown', quantity -> 1, discount -> 0).
  - Validate numeric ranges; abort on undecodable values rather than guessing.
  - Add Year/Month/Month_Start/Year_Month, Shipping_Days, Profit_Margin, flags.

Run from project root: python src/02_clean.py
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = PROJECT_ROOT / "data" / "raw"
CLEAN_DIR = PROJECT_ROOT / "data" / "clean"
REAL_PATH = RAW_DIR / "superstore.csv"
SYNTHETIC_PATH = RAW_DIR / "synthetic_superstore.csv"
CLEAN_PATH = CLEAN_DIR / "sales_clean.csv"
META_PATH = CLEAN_DIR / "metadata.json"

REQUIRED = [
    "Row ID", "Order ID", "Order Date", "Ship Date", "Ship Mode",
    "Customer ID", "Customer Name", "Segment", "City", "State", "Region",
    "Product ID", "Category", "Sub-Category", "Product Name",
    "Sales", "Quantity", "Discount", "Profit",
]
OPTIONAL = ["Country", "Postal Code"]

HIGH_DISCOUNT_THRESHOLD = 0.40  # >= 40% is flagged; business choice, tunable.

DATE_FORMATS = ["%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%d-%b-%Y"]


def _step(n: int, total: int, msg: str) -> None:
    print(f"[02] ({n}/{total}) {msg}")


def _select_input() -> tuple[Path, str]:
    if REAL_PATH.exists():
        return REAL_PATH, "real"
    if SYNTHETIC_PATH.exists():
        return SYNTHETIC_PATH, "synthetic"
    raise FileNotFoundError(
        "No input CSV found. Run src/01_check_data.py for download instructions, "
        "or src/00_generate_sample.py for synthetic data."
    )


def _read_csv(path: Path) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "cp1252"):
        try:
            df = pd.read_csv(path, dtype=str, encoding=enc)
            print(f"     encoding used: {enc}")
            return df
        except UnicodeDecodeError as exc:
            last_err = exc
            continue
    raise RuntimeError(f"could not decode {path}: {last_err}")


def _normalise_headers(df: pd.DataFrame) -> pd.DataFrame:
    # Strip BOM/whitespace and collapse internal spaces so 'Order  ID' == 'Order ID'.
    df.columns = [" ".join(str(c).strip().split()) for c in df.columns]
    return df


def _parse_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value)
    text = str(value).strip()
    if not text:
        return pd.NaT
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return pd.NaT


def _clean_strings(df: pd.DataFrame) -> pd.DataFrame:
    # Collapse whitespace on every text column; standardise casing on the
    # categorical ones so ' consumer ' and 'Consumer' become the same value.
    categorical = {"Segment", "Region", "Category", "Sub-Category", "Ship Mode", "Country", "State"}
    for col in df.columns:
        # At this point every column is still text (numeric conversion happens later),
        # so normalise all of them. (pandas 3 may give dtype 'string', not 'object'.)
        df[col] = df[col].astype("string").str.strip()
        df[col] = df[col].str.replace(r"\s+", " ", regex=True)
        if col in categorical:
            df[col] = df[col].str.title()
    return df


def main() -> int:
    TOTAL = 9
    print("[02] Cleaning sales data...")
    _step(1, TOTAL, "selecting input file")
    try:
        path, source = _select_input()
    except FileNotFoundError as exc:
        print(f"     ERROR: {exc}")
        return 1
    print(f"     source: {source.upper()} ({path.name})")
    if source == "synthetic":
        print("     *** SYNTHETIC DATA - not real business data ***")

    _step(2, TOTAL, "loading CSV")
    try:
        df = _read_csv(path)
    except Exception as exc:  # noqa: BLE001 - report any read failure clearly
        print(f"     ERROR reading file: {exc}")
        return 1
    df = _normalise_headers(df)
    rows_in = len(df)
    print(f"     rows loaded: {rows_in:,}")

    # Required-column check (01 already did this, but re-check for safety).
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        print(f"     ERROR: missing required columns: {missing}")
        return 1
    keep = [c for c in REQUIRED + OPTIONAL if c in df.columns]
    df = df[keep].copy()

    _step(3, TOTAL, "normalising strings + dropping exact duplicates")
    df = _clean_strings(df)
    before = len(df)
    df = df.drop_duplicates(keep="first").reset_index(drop=True)
    duplicates_removed = before - len(df)
    print(f"     duplicates removed: {duplicates_removed:,}  (rows now {len(df):,})")

    _step(4, TOTAL, "parsing dates (explicit formats only)")
    for col in ("Order Date", "Ship Date"):
        before_null = df[col].isna().sum()
        parsed = df[col].map(_parse_date)
        bad = parsed.isna() & df[col].notna()
        if bad.sum() > 0:
            examples = df.loc[bad, col].head(5).tolist()
            print(f"     ERROR: {bad.sum()} unparseable value(s) in '{col}': {examples}")
            print("     Fix the source file or add a format to DATE_FORMATS.")
            return 1
        df[col] = pd.to_datetime(parsed)
        print(f"     {col}: parsed OK ({before_null} originally null)")

    _step(5, TOTAL, "converting numeric columns")
    for col in ("Sales", "Profit", "Discount"):
        original = df[col].copy()
        converted = pd.to_numeric(df[col], errors="coerce")
        bad = converted.isna() & original.notna()
        if bad.sum() > 0:
            print(f"     ERROR: {bad.sum()} non-numeric value(s) in '{col}': {original[bad].head(5).tolist()}")
            return 1
        df[col] = converted
    # Quantity: nullable integer (pandas 'Int64').
    q_converted = pd.to_numeric(df["Quantity"], errors="coerce")
    q_bad = q_converted.isna() & df["Quantity"].notna()
    if q_bad.sum() > 0:
        print(f"     ERROR: {q_bad.sum()} non-numeric value(s) in 'Quantity'")
        return 1
    df["Quantity"] = q_converted.astype("Int64")
    # Row ID: keep as string identifier (never do maths on an ID).
    df["Row ID"] = df["Row ID"].astype("string").str.strip()
    print("     Sales, Profit, Discount, Quantity, Row ID converted")

    _step(6, TOTAL, "handling nulls")
    null_before = df.isna().sum()
    null_cols = null_before[null_before > 0]
    if len(null_cols):
        print("     nulls found:")
        for col, n in null_cols.items():
            print(f"       {col}: {n}")
    else:
        print("     no missing values found")

    critical = ["Order ID", "Order Date", "Ship Date", "Sales", "Profit"]
    before = len(df)
    df = df.dropna(subset=critical).reset_index(drop=True)
    dropped_critical = before - len(df)
    print(f"     dropped rows missing critical fields: {dropped_critical:,}")

    # Optional fills: a missing discount means "no discount recorded" (0),
    # a missing quantity means one unit, missing text becomes 'Unknown'.
    df["Discount"] = df["Discount"].fillna(0.0)
    df["Quantity"] = df["Quantity"].fillna(1).astype("Int64")
    text_cols = ["Customer Name", "Segment", "City", "State", "Region", "Category",
                 "Sub-Category", "Ship Mode", "Product Name", "Product ID",
                 "Customer ID", "Order ID"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna("Unknown")
    if "Country" in df.columns:
        df["Country"] = df["Country"].fillna("Unknown")
    if "Postal Code" in df.columns:
        df["Postal Code"] = df["Postal Code"].fillna("").astype("string")
    print("     optional nulls filled (discount->0, quantity->1, text->'Unknown')")

    # Drop rows that are impossible to analyse (negative/zero sales, bad quantity,
    # discount outside [0,1], ship date before order date). Print every count.
    before = len(df)
    bad_sales = (df["Sales"] <= 0)
    bad_qty = (df["Quantity"] <= 0)
    bad_disc = (df["Discount"] < 0) | (df["Discount"] > 1)
    bad_ship = (df["Ship Date"] < df["Order Date"])
    bad_mask = bad_sales | bad_qty | bad_disc | bad_ship
    print(f"     dropped: sales<=0 {bad_sales.sum()}, qty<=0 {bad_qty.sum()}, "
          f"discount out of [0,1] {bad_disc.sum()}, ship<order {bad_ship.sum()}")
    df = df.loc[~bad_mask].reset_index(drop=True)
    print(f"     rows now {len(df):,} (dropped {bad_mask.sum():,})")

    _step(7, TOTAL, "adding derived columns")
    df["Year"] = df["Order Date"].dt.year.astype("Int64")
    df["Month"] = df["Order Date"].dt.month.astype("Int64")
    df["Month_Name"] = df["Order Date"].dt.strftime("%b %Y")
    df["Month_Start"] = df["Order Date"].dt.to_period("M").dt.to_timestamp()
    df["Year_Month"] = df["Month_Start"].dt.strftime("%Y-%m")
    df["Shipping_Days"] = (df["Ship Date"] - df["Order Date"]).dt.days.astype("Int64")
    # Guard against zero sales (already filtered, but be explicit).
    df["Profit_Margin"] = np.where(df["Sales"] > 0, df["Profit"] / df["Sales"], 0.0)
    df["Profit_Margin"] = df["Profit_Margin"].round(4)
    df["Flag_Negative_Profit"] = (df["Profit"] < 0).astype("Int64")
    df["Flag_High_Discount"] = (df["Discount"] >= HIGH_DISCOUNT_THRESHOLD).astype("Int64")
    print(f"     added Year, Month, Month_Name, Month_Start, Year_Month, "
          f"Shipping_Days, Profit_Margin, Flag_Negative_Profit, Flag_High_Discount")
    print(f"     high-discount threshold = {HIGH_DISCOUNT_THRESHOLD:.0%}")

    _step(8, TOTAL, "validating + saving")
    # No nulls should remain in critical columns.
    if df[["Order ID", "Order Date", "Ship Date", "Sales", "Profit"]].isna().any().any():
        print("     ERROR: nulls remain in critical columns - aborting")
        return 1
    # Keep full source precision for money columns (Sales/Profit can have up to 4
    # decimals in this dataset); rounding each line would shift the totals.
    # Only Discount is a rate, so 2 decimals is plenty. Display rounding happens
    # in Excel/CSV number formats, not in the data.
    df["Discount"] = df["Discount"].round(2)
    # Sort chronologically so the file reads like a ledger.
    df = df.sort_values(["Order Date", "Order ID", "Row ID"], kind="mergesort").reset_index(drop=True)

    CLEAN_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(CLEAN_PATH, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    print(f"     saved: {CLEAN_PATH}  ({len(df):,} rows)")

    metadata = {
        "source": source,
        "raw_file": path.name,
        "rows_in": int(rows_in),
        "rows_out": int(len(df)),
        "duplicates_removed": int(duplicates_removed),
        "rows_dropped_critical_nulls": int(dropped_critical),
        "rows_dropped_invalid": int(bad_mask.sum()),
        "high_discount_threshold": HIGH_DISCOUNT_THRESHOLD,
        "flag_negative_profit": int(df["Flag_Negative_Profit"].sum()),
        "flag_high_discount": int(df["Flag_High_Discount"].sum()),
        "date_min": df["Order Date"].min().strftime("%Y-%m-%d"),
        "date_max": df["Order Date"].max().strftime("%Y-%m-%d"),
        "total_sales": float(df["Sales"].sum()),
        "total_profit": float(df["Profit"].sum()),
        "orders": int(df["Order ID"].nunique()),
    }
    META_PATH.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    print(f"     metadata: {META_PATH}")

    _step(9, TOTAL, "summary")
    print(f"     source        : {source}")
    print(f"     rows in/out   : {rows_in:,} -> {len(df):,}")
    print(f"     duplicates    : {duplicates_removed:,}")
    print(f"     total sales   : {df['Sales'].sum():,.2f}")
    print(f"     total profit  : {df['Profit'].sum():,.2f}")
    print(f"     margin        : {df['Profit'].sum() / df['Sales'].sum():.2%}")
    print(f"     orders        : {df['Order ID'].nunique():,}")
    print(f"     loss lines    : {int((df['Profit'] < 0).sum()):,}")
    print(f"     high-discount : {int((df['Discount'] >= HIGH_DISCOUNT_THRESHOLD).sum()):,}")
    print("[02] Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""Generate a realistic SYNTHETIC sales CSV mirroring the Kaggle Superstore schema.

Why this exists: the official dataset needs a free Kaggle account. This lets a
reviewer (or a client with no account) run the whole pipeline immediately. It is
deliberately labelled SYNTHETIC everywhere downstream so it can never be mistaken
for real business data.

Reproducible: seed=42. Also injects controlled messiness (duplicates, nulls,
stray whitespace) so the cleaning step has something real to do.

Run from project root: python src/00_generate_sample.py
"""
from __future__ import annotations

import calendar
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
N_ORDERS = 600
MIN_LINES, MAX_LINES = 1, 4
START_YEAR, END_YEAR = 2014, 2017

# Q4 peak (Nov/Dec heaviest) is typical for retail.
MONTH_WEIGHTS = np.array([0.07, 0.06, 0.08, 0.08, 0.08, 0.08, 0.08, 0.08, 0.09, 0.09, 0.11, 0.10], dtype=float)
MONTH_WEIGHTS /= MONTH_WEIGHTS.sum()

REGION_STATES = {
    "East": [("New York", "New York"), ("Pennsylvania", "Philadelphia"), ("Massachusetts", "Boston")],
    "Central": [("Illinois", "Chicago"), ("Ohio", "Columbus"), ("Texas", "Houston")],
    "South": [("Florida", "Miami"), ("Georgia", "Atlanta"), ("North Carolina", "Charlotte")],
    "West": [("California", "Los Angeles"), ("Washington", "Seattle"), ("Colorado", "Denver")],
}

SEGMENTS = ["Consumer", "Corporate", "Home Office"]
SEGMENT_WEIGHTS = np.array([0.52, 0.30, 0.18])
SHIP_MODES = ["Standard Class", "Second Class", "First Class", "Same Day"]
SHIP_MODE_WEIGHTS = np.array([0.60, 0.20, 0.15, 0.05])

# (category, subcategory, unit price range, base margin)
# Tables have thin margins (real Superstore pattern); high discounts push profit negative.
PRODUCT_CATALOG = [
    ("Furniture", "Bookcases", (80, 320), 0.06),
    ("Furniture", "Chairs", (120, 520), 0.22),
    ("Furniture", "Tables", (150, 800), 0.03),
    ("Furniture", "Furnishings", (15, 120), 0.30),
    ("Office Supplies", "Binders", (4, 28), 0.38),
    ("Office Supplies", "Paper", (3, 22), 0.42),
    ("Office Supplies", "Storage", (12, 95), 0.25),
    ("Office Supplies", "Art", (8, 60), 0.35),
    ("Office Supplies", "Appliances", (25, 220), 0.20),
    ("Technology", "Phones", (120, 620), 0.30),
    ("Technology", "Accessories", (8, 95), 0.45),
    ("Technology", "Copiers", (400, 1500), 0.35),
    ("Technology", "Machines", (250, 1100), 0.22),
]

DISCOUNT_CHOICES = np.array([0.0, 0.1, 0.15, 0.2, 0.3, 0.4, 0.5, 0.7])
DISCOUNT_WEIGHTS = np.array([0.55, 0.15, 0.10, 0.08, 0.06, 0.04, 0.02, 0.00])  # furniture overrides below

BRANDS = ["Apex", "Northline", "Crest", "Vista", "Orbit", "Summit", "Lumen", "Harbor"]
FIRST_NAMES = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda", "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah", "Charles", "Karen", "Daniel", "Nancy", "Matthew", "Lisa"]
LAST_NAMES = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White"]


def _make_products(rng: np.random.Generator) -> list[dict]:
    products = []
    pid = 1
    for category, sub, (lo, hi), margin in PRODUCT_CATALOG:
        n = rng.integers(4, 9)  # 4-8 products per subcategory
        for _ in range(n):
            brand = rng.choice(BRANDS)
            products.append({
                "Product ID": f"OFF-{category[:3].upper()}-{sub[:3].upper()}-{pid:05d}",
                "Category": category,
                "Sub-Category": sub,
                "Product Name": f"{brand} {sub} - {rng.choice(['Standard', 'Pro', 'Elite', 'Compact', 'XL'])}",
                "price": float(rng.uniform(lo, hi)),
                "margin": float(margin),
            })
            pid += 1
    return products


def _make_customers(rng: np.random.Generator) -> list[dict]:
    customers = []
    for i in range(250):
        region = rng.choice(list(REGION_STATES.keys()))
        state, city = rng.choice(REGION_STATES[region])
        customers.append({
            "Customer ID": f"CG-{10000 + i}",
            "Customer Name": f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}",
            "Segment": rng.choice(SEGMENTS, p=SEGMENT_WEIGHTS),
            "Region": region,
            "State": state,
            "City": city,
        })
    return customers


def _random_order_date(rng: np.random.Generator) -> date:
    year = int(rng.choice(range(START_YEAR, END_YEAR + 1)))
    month = int(rng.choice(range(1, 13), p=MONTH_WEIGHTS))
    day = int(rng.integers(1, calendar.monthrange(year, month)[1] + 1))
    return date(year, month, day)


def _discount_for(sub: str, rng: np.random.Generator) -> float:
    # Furniture (especially Tables) gets deeper discounts in real life.
    if sub in ("Tables", "Bookcases"):
        w = np.array([0.30, 0.18, 0.14, 0.12, 0.12, 0.08, 0.04, 0.02])
    else:
        w = DISCOUNT_WEIGHTS.copy()
    w = w / w.sum()
    return float(rng.choice(DISCOUNT_CHOICES, p=w))


def _ship_days(ship_mode: str, rng: np.random.Generator) -> int:
    if ship_mode == "Same Day":
        return 0
    if ship_mode == "First Class":
        return int(rng.integers(1, 3))
    if ship_mode == "Second Class":
        return int(rng.integers(2, 5))
    return int(rng.integers(4, 8))  # Standard Class


def main() -> None:
    rng = np.random.default_rng(SEED)
    products = _make_products(rng)
    customers = _make_customers(rng)

    rows = []
    order_no = 1
    for _ in range(N_ORDERS):
        cust = customers[int(rng.integers(0, len(customers)))]
        order_date = _random_order_date(rng)
        order_id = f"CA-{order_date.year}-{100000 + order_no}"
        ship_mode = str(rng.choice(SHIP_MODES, p=SHIP_MODE_WEIGHTS))
        ship_date = order_date + timedelta(days=_ship_days(ship_mode, rng))
        n_lines = int(rng.integers(MIN_LINES, MAX_LINES + 1))
        for _ in range(n_lines):
            prod = products[int(rng.integers(0, len(products)))]
            qty = int(rng.integers(1, 10))
            discount = _discount_for(prod["Sub-Category"], rng)
            gross = qty * prod["price"]
            sales = round(gross * (1 - discount), 2)
            # cost uses base margin; discount comes out of margin, so deep discounts lose money
            cost = gross * (1 - prod["margin"]) * float(rng.normal(1.0, 0.03))
            profit = round(sales - cost, 2)
            rows.append({
                "Row ID": len(rows) + 1,
                "Order ID": order_id,
                "Order Date": order_date.strftime("%m/%d/%Y"),
                "Ship Date": ship_date.strftime("%m/%d/%Y"),
                "Ship Mode": ship_mode,
                "Customer ID": cust["Customer ID"],
                "Customer Name": cust["Customer Name"],
                "Segment": cust["Segment"],
                "City": cust["City"],
                "State": cust["State"],
                "Region": cust["Region"],
                "Product ID": prod["Product ID"],
                "Category": prod["Category"],
                "Sub-Category": prod["Sub-Category"],
                "Product Name": prod["Product Name"],
                "Sales": sales,
                "Quantity": qty,
                "Discount": discount,
                "Profit": profit,
            })
        order_no += 1

    df = pd.DataFrame(rows)

    # --- controlled messiness (so the cleaner has real work to do) ---
    # 6 missing Customer Name
    for i in rng.choice(len(df), size=6, replace=False):
        df.loc[i, "Customer Name"] = None
    # 3 missing Quantity
    for i in rng.choice(len(df), size=3, replace=False):
        df.loc[i, "Quantity"] = None
    # 4 missing Discount (cleaner fills with 0 = "no discount recorded")
    for i in rng.choice(len(df), size=4, replace=False):
        df.loc[i, "Discount"] = None
    # 2 missing Sales (critical -> cleaner drops the row)
    for i in rng.choice(len(df), size=2, replace=False):
        df.loc[i, "Sales"] = None
    # 2 missing Profit (critical -> cleaner drops the row)
    for i in rng.choice(len(df), size=2, replace=False):
        df.loc[i, "Profit"] = None
    # 12 stray whitespace / case variants (cleaner normalises)
    for i in rng.choice(len(df), size=12, replace=False):
        df.loc[i, "Segment"] = "  " + str(df.loc[i, "Segment"]).lower() + "  "
    for i in rng.choice(len(df), size=12, replace=False):
        df.loc[i, "Region"] = "  " + str(df.loc[i, "Region"]).lower() + "  "

    # Assign Row ID now, so duplicate rows keep the SAME Row ID as their original
    # (otherwise they would not be exact duplicates and drop_duplicates would miss them).
    df = df.reset_index(drop=True)
    df["Row ID"] = range(1, len(df) + 1)

    # 10 exact duplicates, injected LAST so they stay exact (after all messiness).
    dup_idx = rng.choice(len(df), size=10, replace=False)
    df = pd.concat([df, df.iloc[dup_idx].copy()], ignore_index=True).reset_index(drop=True)

    out_dir = Path(__file__).resolve().parents[1] / "data" / "raw"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "synthetic_superstore.csv"
    df.to_csv(out_path, index=False, encoding="utf-8")

    print("=" * 60)
    print("SYNTHETIC DATA GENERATED - seed=42 - NOT real business data")
    print("=" * 60)
    print(f"Rows written : {len(df)}")
    print(f"Columns      : {len(df.columns)}")
    _dates = pd.to_datetime(df["Order Date"], format="%m/%d/%Y")
    print(f"Date range   : {_dates.min():%Y-%m-%d} .. {_dates.max():%Y-%m-%d}")
    print(f"Output       : {out_path}")
    print("Injected messiness: 10 duplicate rows, 6 null names, 3 null quantities,")
    print("                    4 null discounts, 2 null sales, 2 null profits,")
    print("                    24 whitespace/case variants.")
    print("Every downstream sheet, chart and PNG will be labelled SYNTHETIC.")


if __name__ == "__main__":
    raise SystemExit(main())

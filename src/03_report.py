# -*- coding: utf-8 -*-
"""Build the Excel dashboard and the README PNG charts from the clean CSV.

Why openpyxl (not pandas.ExcelWriter): we need native Excel charts (so the owner
can hover, filter and read values), conditional formatting, KPI cards and print
setup. pandas only writes cells.

Why values and not formulas: openpyxl never calculates, so a formula cell would
show as 0/blank until Excel recalculates. Values are portable and always right.

Run from project root: python src/03_report.py
"""
from __future__ import annotations

import json
import re
import zipfile
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CLEAN_DIR = PROJECT_ROOT / "data" / "clean"
CLEAN_PATH = CLEAN_DIR / "sales_clean.csv"
META_PATH = CLEAN_DIR / "metadata.json"
OUTPUT_DIR = PROJECT_ROOT / "output"
XLSX_PATH = OUTPUT_DIR / "Sales_Dashboard.xlsx"
FINDINGS_PATH = OUTPUT_DIR / "key_findings.md"

# ---- styling constants (one place to retune the look) -----------------------
NAVY = "1F2A44"
TEAL = "2A9D8F"
AMBER = "E9C46A"
CORAL = "E76F51"
GREEN = "067647"
RED = "B42318"
LIGHT_TEAL = "E6F4F1"
LIGHT_GRAY = "F3F4F6"
BORDER = "D1D5DB"
TEXT = "111827"
MUTED = "6B7280"

CURRENCY = "$"
CURRENCY_NAME = "USD"

# Thin borders and fills, reused everywhere.
THIN = Side(style="thin", color=BORDER)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_HEADER = PatternFill("solid", fgColor=NAVY)
FILL_CARD = PatternFill("solid", fgColor=LIGHT_TEAL)
FILL_TOTAL = PatternFill("solid", fgColor=LIGHT_GRAY)
FILL_SYNTH = PatternFill("solid", fgColor="FDF0D5")
FILL_REAL = PatternFill("solid", fgColor=LIGHT_TEAL)
FONT_TITLE = Font(name="Calibri", size=18, bold=True, color=NAVY)
FONT_SUB = Font(name="Calibri", size=10, color=MUTED, italic=True)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=10, color=TEXT)
FONT_SMALL = Font(name="Calibri", size=9, color=MUTED)
FONT_KPI_LABEL = Font(name="Calibri", size=10, color=MUTED)
FONT_KPI_VALUE = Font(name="Calibri", size=20, bold=True, color=NAVY)
FONT_KPI_NOTE = Font(name="Calibri", size=8, color=MUTED, italic=True)


def _money_fmt(decimals: int = 0) -> str:
    return f'"{CURRENCY}"#,##0' + (".00" if decimals else "")


def _step(n: int, total: int, msg: str) -> None:
    print(f"[03] ({n}/{total}) {msg}")


# ---------------------------------------------------------------------------
# Data loading + aggregation
# ---------------------------------------------------------------------------
def load_clean_data():
    if not CLEAN_PATH.exists():
        raise FileNotFoundError(
            f"Clean data not found: {CLEAN_PATH}\n"
            "Run src/01_check_data.py then src/02_clean.py first."
        )
    df = pd.read_csv(CLEAN_PATH, parse_dates=["Order Date", "Ship Date", "Month_Start"])
    if df.empty:
        raise ValueError("clean CSV has no rows")
    source = "unknown"
    metadata = {}
    if META_PATH.exists():
        try:
            metadata = json.loads(META_PATH.read_text(encoding="utf-8"))
            source = metadata.get("source", "unknown")
        except (json.JSONDecodeError, OSError):
            print("     WARNING: metadata.json unreadable; source labelled 'unknown'")
    else:
        print("     WARNING: metadata.json missing; run 02 again to regenerate it")
    # Sanity: metadata row count should match the file.
    if metadata.get("rows_out") is not None and int(metadata["rows_out"]) != len(df):
        raise ValueError(
            f"metadata rows_out ({metadata['rows_out']}) != CSV rows ({len(df)}). "
            "Rerun src/02_clean.py."
        )
    # Required derived columns.
    for col in ("Year", "Month", "Month_Start", "Year_Month", "Profit_Margin",
                "Flag_Negative_Profit", "Flag_High_Discount"):
        if col not in df.columns:
            raise ValueError(f"clean CSV missing derived column '{col}' - rerun 02")
    if df["Sales"].sum() <= 0:
        raise ValueError("total sales is zero; cannot compute margin or AOV")
    return df, source, metadata


def build_aggregates(df: pd.DataFrame):
    total_sales = float(df["Sales"].sum())
    total_profit = float(df["Profit"].sum())
    total_margin = total_profit / total_sales
    orders = int(df["Order ID"].nunique())
    customers = int(df["Customer ID"].nunique())
    units = int(df["Quantity"].sum())
    lines = int(len(df))
    aov = total_sales / orders
    loss_lines = int((df["Profit"] < 0).sum())
    loss_amount = float(df.loc[df["Profit"] < 0, "Profit"].sum())
    high_disc_lines = int((df["Discount"] >= 0.4).sum())
    high_disc_sales = float(df.loc[df["Discount"] >= 0.4, "Sales"].sum())
    high_disc_loss = float(df.loc[(df["Discount"] >= 0.4) & (df["Profit"] < 0), "Profit"].sum())

    # Category
    cat = (df.groupby("Category", as_index=False)
           .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                Orders=("Order ID", "nunique"), Units=("Quantity", "sum"),
                AvgDiscount=("Discount", "mean"))
           .sort_values("Sales", ascending=False).reset_index(drop=True))
    cat["Margin"] = cat["Profit"] / cat["Sales"]
    cat["Share"] = cat["Sales"] / total_sales

    # Sub-category
    sub = (df.groupby(["Category", "Sub-Category"], as_index=False)
           .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                Units=("Quantity", "sum"),
                LossLines=("Profit", lambda s: int((s < 0).sum())))
           .sort_values("Profit", ascending=True).reset_index(drop=True))
    sub["Margin"] = sub["Profit"] / sub["Sales"]

    # Region
    reg = (df.groupby("Region", as_index=False)
           .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                Orders=("Order ID", "nunique"), Customers=("Customer ID", "nunique"))
           .sort_values("Sales", ascending=False).reset_index(drop=True))
    reg["Margin"] = reg["Profit"] / reg["Sales"]
    reg["Share"] = reg["Sales"] / total_sales

    # Monthly (fill missing months with zeros so gaps are visible, not hidden)
    monthly = (df.groupby(["Month_Start", "Year_Month"], as_index=False)
               .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                    Orders=("Order ID", "nunique"))
               .sort_values("Month_Start").reset_index(drop=True))
    full_months = pd.date_range(monthly["Month_Start"].min(),
                               monthly["Month_Start"].max(), freq="MS")
    monthly = (monthly.set_index("Month_Start")
               .reindex(full_months).reset_index()
               .rename(columns={"index": "Month_Start"}))
    monthly["Year_Month"] = monthly["Month_Start"].dt.strftime("%Y-%m")
    monthly["Sales"] = monthly["Sales"].fillna(0.0)
    monthly["Profit"] = monthly["Profit"].fillna(0.0)
    monthly["Orders"] = monthly["Orders"].fillna(0).astype(int)
    monthly["Margin"] = np.where(monthly["Sales"] > 0, monthly["Profit"] / monthly["Sales"], 0.0)
    monthly["AOV"] = np.where(monthly["Orders"] > 0, monthly["Sales"] / monthly["Orders"], 0.0)
    filled_months = int(((monthly["Sales"] == 0) & (monthly["Orders"] == 0)).sum())
    # ^ counts months with no sales AND no orders (true gaps)

    # Top products
    prod = (df.groupby(["Product ID", "Product Name", "Category", "Sub-Category"], as_index=False)
            .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                 Units=("Quantity", "sum"), Orders=("Order ID", "nunique"))
            .sort_values("Sales", ascending=False).reset_index(drop=True))
    prod["Margin"] = prod["Profit"] / prod["Sales"]
    top_n = prod.head(15).copy()
    top_n.insert(0, "Rank", range(1, len(top_n) + 1))
    loss_products = prod[prod["Profit"] < 0].sort_values("Profit").head(10).copy()

    agg = dict(
        total_sales=total_sales, total_profit=total_profit, total_margin=total_margin,
        orders=orders, customers=customers, units=units, lines=lines, aov=aov,
        loss_lines=loss_lines, loss_amount=loss_amount,
        high_disc_lines=high_disc_lines, high_disc_sales=high_disc_sales,
        high_disc_loss=high_disc_loss, cat=cat, sub=sub, reg=reg, monthly=monthly,
        prod=prod, top_n=top_n, loss_products=loss_products,
        filled_months=filled_months,
        date_min=df["Order Date"].min(), date_max=df["Order Date"].max(),
    )
    return agg


def compute_findings(df, agg, source):
    """Turn aggregates into plain-English findings and actions. All derived."""
    total_sales = agg["total_sales"]
    total_profit = agg["total_profit"]
    total_margin = agg["total_margin"]
    cat = agg["cat"]
    reg = agg["reg"]
    sub = agg["sub"]
    monthly = agg["monthly"]
    prod = agg["prod"]

    top_cat = cat.iloc[0]
    best_cat = cat.sort_values("Margin", ascending=False).iloc[0]
    best_reg = reg.sort_values("Margin", ascending=False).iloc[0]
    worst_reg = reg.sort_values("Margin").iloc[0]
    peak_row = monthly.loc[monthly["Sales"].idxmax()]
    peak_label = peak_row["Month_Start"].strftime("%b %Y")
    q4_share = float(monthly.loc[monthly["Month_Start"].dt.month.isin([11, 12]), "Sales"].sum() / total_sales)
    top10_share = float(prod.head(10)["Sales"].sum() / total_sales)
    disc_share = agg["high_disc_sales"] / total_sales if total_sales else 0.0

    # Worst (most negative) sub-category by total profit.
    worst_sub = sub.iloc[0]  # sub is sorted by Profit ascending
    worst_sub_loss = float(worst_sub["Profit"])

    findings = []
    findings.append(
        f"Sales totalled {CURRENCY}{total_sales:,.0f} and profit {CURRENCY}{total_profit:,.0f} "
        f"({total_margin:.1%} margin) across {agg['orders']:,} orders and {agg['lines']:,} line items."
    )
    findings.append(
        f"{top_cat['Category']} is the largest category at {CURRENCY}{top_cat['Sales']:,.0f} "
        f"({top_cat['Share']:.1%} of sales) with a {top_cat['Margin']:.1%} margin."
    )
    findings.append(
        f"{best_reg['Region']} has the best profit margin ({best_reg['Margin']:.1%}); "
        f"{worst_reg['Region']} has the lowest ({worst_reg['Margin']:.1%})."
    )
    if agg["high_disc_lines"] > 0:
        findings.append(
            f"Discounts of 40% or more cover {agg['high_disc_lines']:,} line items "
            f"({disc_share:.1%} of sales) and lose {CURRENCY}{agg['high_disc_loss']:,.0f}."
        )
    findings.append(
        f"Peak month was {peak_label} ({CURRENCY}{peak_row['Sales']:,.0f}); "
        f"Nov-Dec account for {q4_share:.1%} of annual sales."
    )
    if agg["loss_lines"] > 0:
        findings.append(
            f"{agg['loss_lines']:,} line items ({agg['loss_lines']/agg['lines']:.1%} of all) "
            f"are loss-making, totalling {CURRENCY}{agg['loss_amount']:,.0f}."
        )
    if top10_share >= 0.15:
        findings.append(
            f"The top 10 products account for {top10_share:.1%} of all sales "
            f"(concentration risk worth monitoring)."
        )

    # Build up to 3 actions, prioritised by money at stake.
    actions = []
    if worst_sub_loss < 0:
        actions.append(
            f"Cap discounts in {worst_sub['Sub-Category']}: it lost "
            f"{CURRENCY}{abs(worst_sub_loss):,.0f} on {CURRENCY}{worst_sub['Sales']:,.0f} of sales."
        )
    if agg["high_disc_loss"] < 0:
        actions.append(
            f"Require approval for discounts of 40% or more "
            f"({agg['high_disc_lines']:,} lines lose {CURRENCY}{abs(agg['high_disc_loss']):,.0f})."
        )
    if total_margin < 0.12:
        actions.append(
            f"Protect margin: overall margin is {total_margin:.1%}. Push "
            f"{best_cat['Category']} (best margin {best_cat['Margin']:.1%}) in next month's offers."
        )
    else:
        actions.append(
            f"Grow {top_cat['Category']}: it leads sales at {top_cat['Margin']:.1%} margin."
        )
    if q4_share >= 0.28:
        actions.append(
            f"Plan Q4 early: Nov-Dec are {q4_share:.1%} of sales - stock and staff ahead of {peak_label}."
        )
    if len(actions) < 3:
        actions.append(
            f"Target {best_reg['Region']}: highest margin region at {best_reg['Margin']:.1%}."
        )
    actions = actions[:3]
    if not actions:
        actions = ["No urgent action detected - review the monthly trend for seasonality."]

    return findings, actions


def _is_synthetic(source: str) -> bool:
    return source == "synthetic"


def _source_label(source: str) -> str:
    if source == "synthetic":
        return "SYNTHETIC DEMO DATA (seed 42) - not real business data"
    if source == "real":
        return "Source: Kaggle Superstore sample (9,994 line items, 2014-2017) - USD"
    return "Source not recorded - rerun src/02_clean.py"


def _source_fill(source: str):
    if source == "synthetic":
        return FILL_SYNTH
    if source == "real":
        return FILL_REAL
    return PatternFill("solid", fgColor=LIGHT_GRAY)
# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------
def _title_block(ws, title, subtitle, ncols, source, period):
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    for c in range(1, ncols + 1):
        ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    ws.row_dimensions[1].height = 26
    ws["A2"] = subtitle
    ws["A2"].font = FONT_SUB
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.row_dimensions[2].height = 16
    ws["A3"] = _source_label(source) + "  |  Period: " + period
    ws["A3"].font = Font(name="Calibri", size=9, bold=True, color=TEXT)
    ws["A3"].fill = _source_fill(source)
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    ws.row_dimensions[3].height = 18


def _write_table(ws, start_row, headers, rows, formats=None, widths=None,
                 total_row=None, zebra=True, cf_ranges=None):
    formats = formats or {}
    hr = start_row
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=hr, column=j, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    ws.row_dimensions[hr].height = 30
    first = hr + 1
    for i, row in enumerate(rows):
        r = first + i
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = FONT_BODY
            cell.border = BOX
            cell.alignment = Alignment(vertical="center",
                                      horizontal="right" if j > 1 else "left",
                                      wrap_text=(j == 2))
            if zebra and i % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
            fmt = formats.get(j)
            if fmt:
                cell.number_format = fmt
        ws.row_dimensions[r].height = 16
    last_data = first + len(rows) - 1 if rows else first - 1
    if total_row is not None and rows:
        tr = last_data + 1
        for j, val in enumerate(total_row, start=1):
            cell = ws.cell(row=tr, column=j, value=val)
            cell.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
            cell.fill = FILL_TOTAL
            cell.border = Border(top=Side(style="medium", color=NAVY),
                                 left=THIN, right=THIN, bottom=THIN)
            cell.alignment = Alignment(vertical="center",
                                      horizontal="right" if j > 1 else "left")
            fmt = formats.get(j)
            if fmt:
                cell.number_format = fmt
        ws.row_dimensions[tr].height = 18
        last_data = tr
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return hr, first, last_data


def _sheet_setup(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.right.text = "Sales Dashboard"


def _merge_style(ws, rng, value, font, fill=None, alignment=None, border=None):
    ws.merge_cells(rng)
    for row in ws[rng]:
        for cell in row:
            cell.font = font
            if fill is not None:
                cell.fill = fill
            if alignment is not None:
                cell.alignment = alignment
            if border is not None:
                cell.border = border
    top_left = ws[rng.split(":")[0]]
    top_left.value = value
    return top_left


# ---------------------------------------------------------------------------
# Sheet: Summary
# ---------------------------------------------------------------------------
def write_summary(wb, df, agg, source, period, findings, actions):
    ws = wb.create_sheet("Summary")
    ws.sheet_properties.tabColor = NAVY
    _sheet_setup(ws)
    for col, w in zip("ABCDEFG", [3, 20, 20, 20, 20, 20, 20]):
        ws.column_dimensions[col].width = w
    _title_block(ws, "Sales Performance Dashboard",
                 "KPIs, findings and actions for a small-business owner", 7, source, period)

    kpis = [
        ("Total sales", agg["total_sales"], _money_fmt(0), "All line items, after discounts"),
        ("Total profit", agg["total_profit"], _money_fmt(0), "Sales minus cost"),
        ("Profit margin", agg["total_margin"], "0.0%", "Profit per $1 of sales"),
        ("Orders", agg["orders"], "#,##0", "Unique order IDs"),
        ("Avg order value", agg["aov"], _money_fmt(2), "Sales / orders"),
        ("Loss-making lines", agg["loss_lines"] / agg["lines"] if agg["lines"] else 0, "0.0%",
         f"{agg['loss_lines']:,} of {agg['lines']:,} line items"),
    ]
    for j, (label, value, fmt, note) in enumerate(kpis, start=2):
        c = ws.cell(row=5, column=j, value=label)
        c.font = FONT_KPI_LABEL
        c.fill = FILL_CARD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        v = ws.cell(row=6, column=j, value=value)
        v.font = FONT_KPI_VALUE
        v.fill = FILL_CARD
        v.number_format = fmt
        v.alignment = Alignment(horizontal="center", vertical="center")
        v.border = BOX
        n = ws.cell(row=7, column=j, value=note)
        n.font = FONT_KPI_NOTE
        n.fill = FILL_CARD
        n.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        n.border = BOX
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 26

    _merge_style(ws, "B9:G9",
                 "How to read: Sales = money collected. Profit = what is left after cost. "
                 "Margin = profit per $1 of sales (higher is better). AOV = average order value.",
                 Font(name="Calibri", size=9, color=TEXT), fill=PatternFill("solid", fgColor=LIGHT_GRAY),
                 alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
    ws.row_dimensions[9].height = 26

    ws["B11"] = "What to do next"
    ws["B11"].font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells("B11:G11")
    for c in range(2, 8):
        ws.cell(row=11, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    for i, act in enumerate(actions):
        r = 12 + i
        _merge_style(ws, f"B{r}:G{r}", f"{i + 1}.  {act}",
                     Font(name="Calibri", size=10, bold=True, color=TEXT),
                     fill=FILL_CARD,
                     alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
        ws.row_dimensions[r].height = max(28, 14 * (len(act) // 95 + 1))

    hr = 12 + len(actions) + 1
    ws.cell(row=hr, column=2, value="What the data says").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=hr, start_column=2, end_row=hr, end_column=7)
    for c in range(2, 8):
        ws.cell(row=hr, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    for i, fnd in enumerate(findings[:4]):
        r = hr + 1 + i
        _merge_style(ws, f"B{r}:G{r}", "-  " + fnd, FONT_BODY,
                     alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
        ws.row_dimensions[r].height = max(26, 13 * (len(fnd) // 95 + 1))

    flags_row = hr + 1 + min(4, len(findings)) + 1
    flags_text = ("Flags: "
                  f"{agg['loss_lines']:,} loss-making line items ({agg['loss_lines'] / agg['lines']:.1%}); "
                  f"{agg['high_disc_lines']:,} high-discount lines (>=40%); "
                  f"{agg['filled_months']} months with no sales in range; "
                  f"{agg['orders']:,} orders, {agg['customers']:,} customers, {agg['units']:,} units.")
    _merge_style(ws, f"B{flags_row}:G{flags_row}", flags_text, FONT_SMALL,
                 alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
    ws.row_dimensions[flags_row].height = 24

    chart1_row = flags_row + 2
    ws.cell(row=chart1_row, column=2, value="Monthly sales").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=chart1_row, start_column=2, end_row=chart1_row, end_column=7)
    c1 = LineChart()
    c1.title = "Monthly sales"
    c1.height = 9
    c1.width = 17.5
    c1.y_axis.title = "Sales (USD)"
    c1.y_axis.numFmt = _money_fmt(0)
    c1.x_axis.title = "Month"
    c1.legend.position = "b"
    data = Reference(wb["Monthly_Trend"], min_col=2, min_row=25, max_row=25 + len(agg["monthly"]) - 1)
    cats = Reference(wb["Monthly_Trend"], min_col=1, min_row=26, max_row=25 + len(agg["monthly"]))
    c1.add_data(data, titles_from_data=True)
    c1.set_categories(cats)
    for s in c1.series:
        s.graphicalProperties.line.solidFill = TEAL
        s.graphicalProperties.line.width = 18000
        s.marker.symbol = "circle"
        s.marker.size = 4
        s.marker.graphicalProperties.solidFill = NAVY
    ws.add_chart(c1, f"B{chart1_row + 1}")

    chart2_row = chart1_row + 20
    ws.cell(row=chart2_row, column=2, value="Sales by category").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=chart2_row, start_column=2, end_row=chart2_row, end_column=7)
    c2 = BarChart()
    c2.type = "col"
    c2.title = "Sales by category"
    c2.height = 9
    c2.width = 17.5
    c2.y_axis.title = "Sales (USD)"
    c2.y_axis.numFmt = _money_fmt(0)
    c2.legend = None
    data2 = Reference(wb["Category"], min_col=2, min_row=4, max_row=4 + len(agg["cat"]))
    cats2 = Reference(wb["Category"], min_col=1, min_row=5, max_row=4 + len(agg["cat"]))
    c2.add_data(data2, titles_from_data=True)
    c2.set_categories(cats2)
    for s in c2.series:
        s.graphicalProperties.solidFill = NAVY
    ws.add_chart(c2, f"B{chart2_row + 1}")

    footer_row = chart2_row + 20
    footer = ("Rows in: {rin:,}  |  Rows out: {rout:,}  |  Duplicates removed: {dup:,}  |  "
              "Negative-profit lines: {neg:,}  |  High-discount lines: {hd:,}  |  "
              "Generated: {gen}  |  Source: {src}").format(
        rin=0, rout=len(df), dup=0, neg=agg["loss_lines"], hd=agg["high_disc_lines"],
        gen=datetime.now().strftime("%Y-%m-%d %H:%M"), src=source)
    try:
        meta = json.loads(META_PATH.read_text(encoding="utf-8"))
        footer = footer.replace("Rows in: 0", f"Rows in: {meta.get('rows_in', 0):,}")
        footer = footer.replace("Duplicates removed: 0", f"Duplicates removed: {meta.get('duplicates_removed', 0):,}")
    except (OSError, json.JSONDecodeError):
        pass
    _merge_style(ws, f"B{footer_row}:G{footer_row}", footer, FONT_SMALL,
                 alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
    ws.row_dimensions[footer_row].height = 28


# ---------------------------------------------------------------------------
# Sheet: Category
# ---------------------------------------------------------------------------
def write_category(wb, agg, source, period):
    ws = wb.create_sheet("Category")
    ws.sheet_properties.tabColor = TEAL
    _sheet_setup(ws)
    widths = [26, 14, 14, 10, 10, 10, 10]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    dfc = agg["cat"]
    _title_block(ws, "Category performance",
                 "Sales, profit and margin by category and sub-category", 7,
                 source, period)

    cat_rows = []
    for _, r in dfc.iterrows():
        cat_rows.append([r["Category"], r["Sales"], r["Profit"], r["Margin"],
                         r["Orders"], r["Units"], r["Share"]])
    total_row = ["Total", dfc["Sales"].sum(), dfc["Profit"].sum(),
                 dfc["Profit"].sum() / dfc["Sales"].sum(), dfc["Orders"].sum(),
                 dfc["Units"].sum(), 1.0]
    fmts = {2: _money_fmt(0), 3: _money_fmt(0), 4: "0.0%", 5: "#,##0", 6: "#,##0", 7: "0.0%"}
    hr, first, last = _write_table(ws, 4, ["Category", "Sales", "Profit", "Margin",
                                          "Orders", "Units", "Share of sales"],
                                  cat_rows, formats=fmts, total_row=total_row)
    ws.freeze_panes = "A5"

    sub_start = last + 3
    ws.cell(row=sub_start, column=1, value="Sub-category detail (sorted by profit, lowest first)").font = \
        Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=sub_start, start_column=1, end_row=sub_start, end_column=7)
    for c in range(1, 8):
        ws.cell(row=sub_start, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)

    sub_rows = []
    for _, r in agg["sub"].iterrows():
        sub_rows.append([f"{r['Category']}  |  {r['Sub-Category']}", r["Sales"], r["Profit"],
                         r["Margin"], r["Units"], r["LossLines"], None])
    sub_total = ["Total", agg["sub"]["Sales"].sum(), agg["sub"]["Profit"].sum(),
                 agg["sub"]["Profit"].sum() / agg["sub"]["Sales"].sum(),
                 agg["sub"]["Units"].sum(), agg["sub"]["LossLines"].sum(), None]
    sub_fmts = {2: _money_fmt(0), 3: _money_fmt(0), 4: "0.0%", 5: "#,##0", 6: "#,##0"}
    _write_table(ws, sub_start + 1,
                 ["Sub-category", "Sales", "Profit", "Margin", "Units", "Loss-making lines", ""],
                 sub_rows, formats=sub_fmts, total_row=sub_total)
    sub_first = sub_start + 2
    sub_last = sub_first + len(sub_rows) - 1

    chart_row = sub_last + 3
    ws.cell(row=chart_row, column=1, value="Profit by sub-category (red = loss)").font = \
        Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=chart_row, start_column=1, end_row=chart_row, end_column=7)
    for c in range(1, 8):
        ws.cell(row=chart_row, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Profit by sub-category"
    ch.height = 16
    ch.width = 17.5
    ch.y_axis.title = "Profit (USD)"
    ch.y_axis.numFmt = _money_fmt(0)
    ch.legend = None
    data = Reference(ws, min_col=3, min_row=sub_start + 1, max_row=sub_last)
    cats = Reference(ws, min_col=1, min_row=sub_start + 2, max_row=sub_last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    from openpyxl.chart.marker import DataPoint
    profits = agg["sub"]["Profit"].tolist()
    for idx, p in enumerate(profits):
        dp = DataPoint(idx=idx)
        dp.graphicalProperties.solidFill = CORAL if p < 0 else TEAL
        ch.series[0].dPt.append(dp)
    ws.add_chart(ch, f"A{chart_row + 1}")

    note_row = chart_row + 33
    _merge_style(ws, f"A{note_row}:G{note_row}",
                 "Sorted by profit (lowest first). Red bars are loss-making sub-categories - "
                 "usually a discount or cost problem, not a demand problem.",
                 FONT_SMALL, alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))


# ---------------------------------------------------------------------------
# Sheet: Region
# ---------------------------------------------------------------------------
def write_region(wb, agg, source, period):
    ws = wb.create_sheet("Region")
    ws.sheet_properties.tabColor = AMBER
    _sheet_setup(ws)
    widths = [16, 16, 14, 12, 10, 12, 10]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    dfr = agg["reg"]
    _title_block(ws, "Region performance", "Sales, profit and margin by region", 7,
                 source, period)

    rows = []
    for _, r in dfr.iterrows():
        rows.append([r["Region"], r["Sales"], r["Profit"], r["Margin"],
                     r["Orders"], r["Customers"], r["Share"]])
    total_row = ["Total", dfr["Sales"].sum(), dfr["Profit"].sum(),
                 dfr["Profit"].sum() / dfr["Sales"].sum(), dfr["Orders"].sum(),
                 dfr["Customers"].sum(), 1.0]
    fmts = {2: _money_fmt(0), 3: _money_fmt(0), 4: "0.0%", 5: "#,##0", 6: "#,##0", 7: "0.0%"}
    hr, first, last = _write_table(ws, 4, ["Region", "Sales", "Profit", "Margin",
                                          "Orders", "Customers", "Share of sales"],
                                  rows, formats=fmts, total_row=total_row)
    ws.freeze_panes = "A5"

    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
    margin_range = f"D5:D{max(5, last - 1)}"
    ws.conditional_formatting.add(margin_range, ColorScaleRule(
        start_type="num", start_value=0, start_color="F4C7C3",
        mid_type="num", mid_value=0.15, mid_color="FDF3D0",
        end_type="num", end_value=0.30, end_color="C6E7D8"))
    ws.conditional_formatting.add(f"C5:C{max(5, last - 1)}", CellIsRule(
        operator="lessThan", formula=["0"],
        font=Font(name="Calibri", size=10, bold=True, color=RED)))
    ws.conditional_formatting.add(f"B5:B{max(5, last - 1)}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color=TEAL, showValue=True))

    chart_row = last + 3
    ws.cell(row=chart_row, column=1, value="Share of sales by region").font = \
        Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=chart_row, start_column=1, end_row=chart_row, end_column=7)
    for c in range(1, 8):
        ws.cell(row=chart_row, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    pie = PieChart()
    pie.title = "Share of sales by region"
    pie.height = 9
    pie.width = 11
    pie.legend.position = "r"
    data = Reference(ws, min_col=2, min_row=4, max_row=4 + len(dfr))
    cats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(dfr))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showCatName = True
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.numFmt = "0.0%"
    from openpyxl.chart.marker import DataPoint
    palette = [NAVY, TEAL, AMBER, CORAL, GREEN]
    for idx in range(len(dfr)):
        dp = DataPoint(idx=idx)
        dp.graphicalProperties.solidFill = palette[idx % len(palette)]
        pie.series[0].dPt.append(dp)
    ws.add_chart(pie, f"A{chart_row + 1}")

    note_row = chart_row + 12
    _merge_style(ws, f"A{note_row}:G{note_row}",
                 "Margin colour scale: red < 10%, amber 10-15%, green 15%+. "
                 "Data bars show relative sales size.",
                 FONT_SMALL, alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
# ---------------------------------------------------------------------------
# Sheet: Monthly_Trend
# ---------------------------------------------------------------------------
def write_monthly(wb, agg, source, period):
    ws = wb.create_sheet("Monthly_Trend")
    ws.sheet_properties.tabColor = "3B82F6"
    _sheet_setup(ws)
    widths = [14, 14, 14, 10, 10, 14]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    m = agg["monthly"]
    _title_block(ws, "Monthly trend", "Sales, profit and margin by month", 6,
                 source, period)

    # Chart first so the trend is visible without scrolling.
    ws.cell(row=4, column=1, value="Monthly profit margin").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=6)
    for c in range(1, 7):
        ws.cell(row=4, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    ch = LineChart()
    ch.title = "Monthly profit margin"
    ch.height = 9
    ch.width = 17.5
    ch.y_axis.title = "Margin"
    ch.y_axis.numFmt = "0.0%"
    ch.x_axis.title = "Month"
    ch.legend = None
    data = Reference(ws, min_col=4, min_row=25, max_row=25 + len(m) - 1)
    cats = Reference(ws, min_col=1, min_row=26, max_row=25 + len(m))
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s in ch.series:
        s.graphicalProperties.line.solidFill = AMBER
        s.graphicalProperties.line.width = 18000
        s.marker.symbol = "circle"
        s.marker.size = 3
        s.marker.graphicalProperties.solidFill = NAVY
    ws.add_chart(ch, "A5")

    # Table below.
    table_row = 25
    ws.cell(row=table_row - 1, column=1, value="Monthly detail").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=table_row - 1, start_column=1, end_row=table_row - 1, end_column=6)
    for c in range(1, 7):
        ws.cell(row=table_row - 1, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    rows = []
    for _, r in m.iterrows():
        rows.append([r["Year_Month"], r["Sales"], r["Profit"], r["Margin"],
                     r["Orders"], r["AOV"]])
    fmts = {2: _money_fmt(0), 3: _money_fmt(0), 4: "0.0%", 5: "#,##0", 6: _money_fmt(2)}
    hr, first, last = _write_table(ws, table_row, ["Month", "Sales", "Profit", "Margin",
                                                  "Orders", "Avg order value"],
                                  rows, formats=fmts)
    ws.freeze_panes = f"A{table_row + 1}"

    note_row = last + 2
    _merge_style(ws, f"A{note_row}:F{note_row}",
                 "Margin = profit / sales. Missing months are shown as zero so a gap is visible "
                 "instead of being hidden by a connecting line.",
                 FONT_SMALL, alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))


# ---------------------------------------------------------------------------
# Sheet: Top_Products
# ---------------------------------------------------------------------------
def write_products(wb, agg, source, period):
    ws = wb.create_sheet("Top_Products")
    ws.sheet_properties.tabColor = CORAL
    _sheet_setup(ws)
    widths = [6, 42, 18, 14, 14, 10, 8, 8]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    _title_block(ws, "Top products", "Highest-selling products and loss-makers to review", 8,
                 source, period)

    top = agg["top_n"]
    rows = []
    for _, r in top.iterrows():
        rows.append([r["Rank"], r["Product Name"], r["Category"], r["Sales"],
                     r["Profit"], r["Margin"], r["Units"], r["Orders"]])
    fmts = {4: _money_fmt(0), 5: _money_fmt(0), 6: "0.0%", 7: "#,##0", 8: "#,##0"}
    hr, first, last = _write_table(ws, 4, ["Rank", "Product", "Category", "Sales",
                                          "Profit", "Margin", "Units", "Orders"],
                                  rows, formats=fmts)
    ws.freeze_panes = "A5"

    # Chart: top 10 by sales (horizontal bar), red where profit is negative.
    chart_row = last + 3
    ws.cell(row=chart_row, column=1, value="Top 10 products by sales").font = \
        Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=chart_row, start_column=1, end_row=chart_row, end_column=8)
    for c in range(1, 9):
        ws.cell(row=chart_row, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    chart_last = min(first + 9, last)
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Top 10 products by sales"
    ch.height = 15
    ch.width = 17.5
    ch.y_axis.title = "Sales (USD)"
    ch.y_axis.numFmt = _money_fmt(0)
    ch.legend = None
    data = Reference(ws, min_col=4, min_row=4, max_row=chart_last)
    cats = Reference(ws, min_col=2, min_row=first, max_row=chart_last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    from openpyxl.chart.marker import DataPoint
    profits = top["Profit"].head(10).tolist()
    for idx, p in enumerate(profits):
        dp = DataPoint(idx=idx)
        dp.graphicalProperties.solidFill = CORAL if p < 0 else NAVY
        ch.series[0].dPt.append(dp)
    ws.add_chart(ch, f"A{chart_row + 1}")

    # Loss-makers table.
    loss_row = chart_row + 33
    ws.cell(row=loss_row, column=1, value="Loss-making products to review").font = \
        Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=loss_row, start_column=1, end_row=loss_row, end_column=8)
    for c in range(1, 9):
        ws.cell(row=loss_row, column=c).fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    loss = agg["loss_products"]
    if len(loss):
        lrows = []
        for _, r in loss.iterrows():
            lrows.append([r["Product Name"], r["Category"], r["Sales"],
                          r["Profit"], r["Margin"], r["Units"]])
        lfmts = {3: _money_fmt(0), 4: _money_fmt(0), 5: "0.0%", 6: "#,##0"}
        _write_table(ws, loss_row + 1, ["Product", "Category", "Sales", "Profit",
                                       "Margin", "Units"], lrows, formats=lfmts)
    else:
        _merge_style(ws, f"A{loss_row + 1}:H{loss_row + 1}",
                     "No loss-making products in this dataset.",
                     FONT_BODY, alignment=Alignment(horizontal="left", vertical="center"))

    note_row = loss_row + (len(loss) + 3 if len(loss) else 3)
    _merge_style(ws, f"A{note_row}:H{note_row}",
                 "Top 15 by sales shown; chart shows top 10. Red bars = negative profit. "
                 "Review discounting and supplier cost on those items.",
                 FONT_SMALL, alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))


# ---------------------------------------------------------------------------
# PNG charts for the README (matplotlib, Agg backend)
# ---------------------------------------------------------------------------
def make_pngs(df, agg, source):
    # matplotlib wants '#RRGGBB'; the module constants are stored without '#'
    # for openpyxl. Snapshot the globals first, then shadow locally.
    _NAVY, _TEAL, _MUTED, _CORAL, _AMBER, _GREEN, _TEXT = (globals()[k] for k in ("NAVY", "TEAL", "MUTED", "CORAL", "AMBER", "GREEN", "TEXT"))
    NAVY = "#" + _NAVY
    TEAL = "#" + _TEAL
    MUTED = "#" + _MUTED
    CORAL = "#" + _CORAL
    AMBER = "#" + _AMBER
    GREEN = "#" + _GREEN
    TEXT = "#" + _TEXT
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    period = f"{agg['date_min']:%b %Y} - {agg['date_max']:%b %Y}"
    src_note = ("Source: synthetic demo data (seed 42) - not real business data"
                if source == "synthetic"
                else "Source: Kaggle Superstore sample (USD)")
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "axes.titlesize": 14, "axes.titleweight": "bold",
        "axes.labelsize": 10, "xtick.labelsize": 9, "ytick.labelsize": 9,
        "axes.spines.top": False, "axes.spines.right": False, "axes.grid": True,
        "grid.color": "#E5E7EB", "grid.linewidth": 0.8, "axes.axisbelow": True,
    })

    def style_fig(fig, title, subtitle):
        fig.text(0.01, 0.98, title, fontsize=15, fontweight="bold", color=NAVY, va="top")
        fig.text(0.01, 0.93, subtitle, fontsize=10, color=MUTED, va="top")
        fig.text(0.01, 0.01, src_note, fontsize=8, color=MUTED, va="bottom")
        fig.subplots_adjust(top=0.86, bottom=0.12, left=0.08, right=0.96)

    # 1. Monthly sales + profit (dual axis)
    fig, ax = plt.subplots(figsize=(11, 6))
    ax2 = ax.twinx()
    months = agg["monthly"]["Month_Start"]
    sales = agg["monthly"]["Sales"]
    profit = agg["monthly"]["Profit"]
    ax.bar(months, sales, width=20, color=TEAL, alpha=0.85, label="Sales")
    ax2.plot(months, profit, color=NAVY, marker="o", markersize=3, linewidth=1.8, label="Profit")
    ax2.axhline(0, color=MUTED, linewidth=0.8, linestyle=":")
    ax.set_ylabel("Sales (USD)")
    ax2.set_ylabel("Profit (USD)")
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{CURRENCY}{x:,.0f}"))
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{CURRENCY}{x:,.0f}"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=6))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    for lbl in ax.get_xticklabels():
        lbl.set_rotation(30)
        lbl.set_ha("right")
    # annotate peak
    peak_idx = int(sales.values.argmax())
    ax.annotate(f"Peak {months.iloc[peak_idx]:%b %Y}\n{CURRENCY}{sales.iloc[peak_idx]:,.0f}",
                xy=(months.iloc[peak_idx], sales.iloc[peak_idx]),
                xytext=(0, 12), textcoords="offset points", ha="center", fontsize=8, color=NAVY,
                arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.8))
    ax.set_title("")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False, fontsize=9)
    style_fig(fig, "Monthly sales and profit", "Bars = sales (left axis); line = profit (right axis); USD.")
    fig.savefig(OUTPUT_DIR / "01_monthly_sales_profit.png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    # 2. Profit by sub-category (horizontal bar)
    sub = agg["sub"].sort_values("Profit", ascending=True)
    fig, ax = plt.subplots(figsize=(11, 7))
    colors = [CORAL if p < 0 else TEAL for p in sub["Profit"]]
    bars = ax.barh(sub["Sub-Category"], sub["Profit"], color=colors, height=0.7)
    ax.axvline(0, color=MUTED, linewidth=0.8)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{CURRENCY}{x:,.0f}"))
    for i, (val, name) in enumerate(zip(sub["Profit"], sub["Sub-Category"])):
        ax.text(val, i, f"  {CURRENCY}{val:,.0f}", va="center",
                ha="left" if val >= 0 else "right", fontsize=8, color=TEXT)
    style_fig(fig, "Profit by sub-category", "Red bars are loss-making. Sorted lowest first.")
    fig.savefig(OUTPUT_DIR / "02_subcategory_profit.png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    # 3. Region share (donut)
    reg = agg["reg"]
    fig, ax = plt.subplots(figsize=(9, 6.5))
    total = reg["Sales"].sum()
    wedges, _ = ax.pie(reg["Sales"], startangle=90, counterclock=False,
                       colors=[NAVY, TEAL, AMBER, CORAL, GREEN][:len(reg)],
                       wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    labels = [f"{r['Region']}   {r['Sales'] / total:.1%}" for _, r in reg.iterrows()]
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False, fontsize=10)
    ax.text(0, 0, f"{CURRENCY}{total / 1e6:.2f}M\nTotal sales", ha="center", va="center",
            fontsize=13, fontweight="bold", color=NAVY)
    ax.set_aspect("equal")
    style_fig(fig, "Share of sales by region", "Donut shows each region's share of total sales.")
    fig.savefig(OUTPUT_DIR / "03_region_share.png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"[03] wrote 3 PNG charts to {OUTPUT_DIR}")
# ---------------------------------------------------------------------------
# Verification + findings file + main
# ---------------------------------------------------------------------------
def write_findings(agg, source, findings, actions):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    lines = []
    lines.append("# Key findings (generated by src/03_report.py)")
    lines.append("")
    lines.append(f"- **Source:** {_source_label(source).replace('Source: ', '', 1)}")
    lines.append(f"- **Period:** {agg['date_min']:%Y-%m-%d} to {agg['date_max']:%Y-%m-%d}")
    lines.append(f"- **Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    lines.append("## KPIs")
    lines.append("")
    lines.append("| Metric | Value |")
    lines.append("|---|---:|")
    lines.append(f"| Total sales | {CURRENCY}{agg['total_sales']:,.2f} |")
    lines.append(f"| Total profit | {CURRENCY}{agg['total_profit']:,.2f} |")
    lines.append(f"| Profit margin | {agg['total_margin']:.2%} |")
    lines.append(f"| Orders | {agg['orders']:,} |")
    lines.append(f"| Average order value | {CURRENCY}{agg['aov']:,.2f} |")
    lines.append(f"| Loss-making line items | {agg['loss_lines']:,} ({agg['loss_lines'] / agg['lines']:.1%}) |")
    lines.append(f"| High-discount line items (>=40%) | {agg['high_disc_lines']:,} |")
    lines.append(f"| Customers | {agg['customers']:,} |")
    lines.append(f"| Units sold | {agg['units']:,} |")
    lines.append("")
    lines.append("## What the data says")
    lines.append("")
    for i, f in enumerate(findings, 1):
        lines.append(f"{i}. {f}")
    lines.append("")
    lines.append("## What to do next")
    lines.append("")
    for i, a in enumerate(actions, 1):
        lines.append(f"{i}. {a}")
    lines.append("")
    lines.append("## Method notes")
    lines.append("")
    lines.append(f"- High-discount flag threshold: {0.4:.0%} (tunable in src/02_clean.py).")
    lines.append("- Negative profit is flagged, not removed: a loss is a business signal.")
    lines.append("- Profit margin = profit / sales, guarded against zero sales.")
    lines.append("- AOV = total sales / unique orders (an order can have many lines).")
    lines.append("- Missing months are filled with zero so gaps stay visible.")
    lines.append("- All figures in USD.")
    FINDINGS_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"[03] wrote findings: {FINDINGS_PATH}")


def verify_workbook(wb, source):
    print("[03] verifying workbook...")
    # Sheet names and order.
    expected_sheets = ["Summary", "Category", "Region", "Monthly_Trend", "Top_Products"]
    actual = wb.sheetnames
    if actual != expected_sheets:
        raise ValueError(f"sheet mismatch: {actual} != {expected_sheets}")
    print(f"     sheets OK: {actual}")

    # Inspect the xlsx zip for native chart XML.
    with zipfile.ZipFile(XLSX_PATH) as z:
        chart_files = [n for n in z.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")]
        types = {"barChart": 0, "lineChart": 0, "pieChart": 0}
        for cf in chart_files:
            xml = z.read(cf).decode("utf-8")
            for t in types:
                if f"<{t}>" in xml or f"<{t} " in xml:
                    types[t] += 1
        print(f"     chart XML files: {len(chart_files)}  -> {types}")
        if len(chart_files) < 4:
            raise ValueError("fewer than 4 charts in workbook")
        if types["barChart"] == 0 or types["lineChart"] == 0 or types["pieChart"] == 0:
            raise ValueError("missing one of bar/line/pie chart types")
        # Every chart must have a title and at least one series.
        for cf in chart_files:
            xml = z.read(cf).decode("utf-8")
            if "<title>" not in xml:
                raise ValueError(f"{cf} has no title")
            if "<ser>" not in xml:
                raise ValueError(f"{cf} has no series")

    # KPI cells on Summary.
    ws = wb["Summary"]
    if ws["B6"].value is None or ws["C6"].value is None:
        raise ValueError("KPI values missing on Summary")
    print(f"     KPIs OK: sales={ws['B6'].value:,.0f} profit={ws['C6'].value:,.0f} margin={ws['D6'].value:.2%}")

    # PNG checks.
    for name in ("01_monthly_sales_profit.png", "02_subcategory_profit.png", "03_region_share.png"):
        p = OUTPUT_DIR / name
        if not p.exists() or p.stat().st_size < 5000:
            raise ValueError(f"PNG missing or too small: {p}")
        # read dimensions via matplotlib (already imported)
        import matplotlib.image as mpimg
        img = mpimg.imread(p)
        h, w = img.shape[:2]
        if w < 800 or h < 400:
            raise ValueError(f"PNG too small dimensions: {p} {w}x{h}")
        print(f"     PNG OK: {name} ({w}x{h}, {p.stat().st_size // 1024} KB)")

    # Findings file.
    if not FINDINGS_PATH.exists() or FINDINGS_PATH.stat().st_size < 100:
        raise ValueError("key_findings.md missing or empty")
    print(f"     findings file OK: {FINDINGS_PATH}")

    # Conditional formatting present.
    cf_count = sum(len(ws2.conditional_formatting._cf_rules) for ws2 in wb.worksheets)
    print(f"     conditional-format rules: {cf_count}")
    if cf_count == 0:
        raise ValueError("no conditional formatting applied")
    print("[03] verification passed")


def main() -> int:
    print("[03] Building dashboard...")
    _step(1, 8, "loading clean data")
    try:
        df, source, metadata = load_clean_data()
    except (FileNotFoundError, ValueError) as exc:
        print(f"     ERROR: {exc}")
        return 1
    print(f"     rows: {len(df):,}  source: {source}")
    if source == "synthetic":
        print("     *** SYNTHETIC DATA - outputs will be labelled synthetic ***")

    _step(2, 8, "building aggregates")
    agg = build_aggregates(df)
    findings, actions = compute_findings(df, agg, source)
    print(f"     categories: {len(agg['cat'])}, sub-categories: {len(agg['sub'])}, "
          f"regions: {len(agg['reg'])}, months: {len(agg['monthly'])}, products: {len(agg['prod'])}")

    _step(3, 8, "creating workbook")
    wb = Workbook()
    # Drop the default blank sheet; each writer creates its own sheet in order.
    wb.remove(wb.active)

    period = f"{agg['date_min']:%b %Y} - {agg['date_max']:%b %Y}"

    _step(4, 8, "writing Category sheet")
    write_category(wb, agg, source, period)
    _step(5, 8, "writing Region sheet")
    write_region(wb, agg, source, period)
    _step(6, 8, "writing Monthly_Trend and Top_Products sheets")
    write_monthly(wb, agg, source, period)
    write_products(wb, agg, source, period)

    _step(7, 8, "writing Summary sheet (KPIs, actions, charts)")
    write_summary(wb, df, agg, source, period, findings, actions)

    # Workbook properties.
    wb.properties.title = "Sales Performance Dashboard"
    wb.properties.subject = "Cleaned sales data: KPIs, category, region, monthly trend, top products"
    wb.properties.creator = "Moaz"
    wb.properties.description = ("Built with pandas + openpyxl + matplotlib. "
                                "Source: " + _source_label(source))
    wb.properties.keywords = "sales; dashboard; openpyxl; pandas"
    # Put Summary first so it is the landing tab.
    if "Summary" in wb.sheetnames and wb.sheetnames.index("Summary") != 0:
        wb.move_sheet("Summary", -(wb.sheetnames.index("Summary")))
    wb.active = 0  # Summary

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(XLSX_PATH)
    except PermissionError:
        print(f"     ERROR: cannot write {XLSX_PATH} - close it in Excel and rerun")
        return 1
    print(f"     saved: {XLSX_PATH}")

    _step(8, 8, "PNG charts + findings file")
    make_pngs(df, agg, source)
    write_findings(agg, source, findings, actions)

    # Verification pass (reopen from disk).
    try:
        wb2 = load_workbook(XLSX_PATH)
    except Exception as exc:  # noqa: BLE001
        print(f"     ERROR: saved workbook cannot be reopened: {exc}")
        return 1
    try:
        verify_workbook(wb2, source)
    except ValueError as exc:
        print(f"     VERIFICATION FAILED: {exc}")
        return 1

    # Print findings for the console transcript.
    print()
    print("=" * 72)
    print("KEY FINDINGS")
    print("=" * 72)
    for i, f in enumerate(findings, 1):
        print(f"  {i}. {f}")
    print()
    print("ACTIONS")
    print("=" * 72)
    for i, a in enumerate(actions, 1):
        print(f"  {i}. {a}")
    print()
    print("[03] Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

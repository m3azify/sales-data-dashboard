# LEARNING.md - how this pipeline works (beginner edition)

This document teaches the project to someone who has never written a data
pipeline. Read it next to the code, not instead of it: open the script, find
the function name, then come back here.

## The big picture

```
raw CSV  ->  01 check  ->  02 clean  ->  03 report  ->  Excel + PNG
```

Each arrow is a separate script with one job. That is the single most important
design decision: if step 2 breaks, steps 1 and 3 are untouched, and you always
know which step failed.

- `00_generate_sample.py` makes fake data so the project runs without a Kaggle
  account. It is NOT part of the pipeline; it is a helper.
- `01_check_data.py` only checks the input. It uses only the standard library so
  it runs before `pip install`.
- `02_clean.py` turns the messy CSV into `data/clean/sales_clean.csv`.
- `03_report.py` turns the clean CSV into the Excel workbook and 3 PNGs.

## Core ideas (with analogies)

- **DataFrame** = a spreadsheet in memory. Rows are records, columns are fields.
  `df["Sales"]` is one column; `df.head()` shows the first 5 rows.
- **dtype** = the type of a column. "123" as text is not the number 123. We
  read everything as text first (`dtype=str`) so pandas cannot guess wrong, then
  convert on purpose.
- **Null / NaN** = a missing cell. Pandas represents it as `NaN` (Not a Number).
  `df.isna().sum()` counts missing values per column.
- **Groupby** = Excel's "PivotTable": split rows into groups, then aggregate
  (sum, mean, count). `df.groupby("Category")["Sales"].sum()`.
- **nunique** = count of distinct values. Orders are counted this way because one
  order has several line items.
- **Margin** = profit / sales. It answers "how much of every dollar stays?".
  Sales alone misleads: a $1M product with 1% margin is worse than a $100k
  product with 40% margin.

## Script 01 - the input gate

- `Path(__file__).resolve().parents[1]` finds the project root no matter which
  folder you run the script from. This is why `python src/01_check_data.py`
  works from the project root.
- It tries `utf-8-sig` then `cp1252`. The first handles a BOM (byte order mark)
  that Excel/Kaggle exports sometimes prepend.
- It validates the header against a required list and aborts with a readable list
  of missing columns. It never downloads anything: no network, no API keys.
- Exit code 1 on failure is deliberate so a parent script or CI can react.

## Script 02 - cleaning (the heart of the project)

Order of operations, and WHY:

1. **Read as text** (`dtype=str`). If pandas guesses types, a column like
   `Postal Code` ("01234") becomes the number 1234 and loses its leading zero.
   Reading as text keeps every value intact; we convert afterwards.
2. **Normalise headers** (strip, collapse spaces). `Order  ID` and `Order ID`
   become the same column.
3. **Normalise strings** (strip, collapse whitespace, title-case only the
   categorical fields). IDs and product names keep their original case.
4. **Drop duplicates** AFTER normalising, so ` Consumer ` and `Consumer` count
   as the same row. We keep the first occurrence.
5. **Parse dates from explicit formats only.** Guessing day vs month silently
   corrupts data (is 03/04/2014 March 4 or April 3?). We try a short list and
   abort if any value fails, printing the bad examples.
6. **Convert numerics with `errors="coerce"`**, then check how many became NaN.
   A value like "N/A" or "$1,000" is reported, not silently turned into 0.
7. **Critical nulls drop the row** (no Order ID / no Sales = no analysis).
   **Optional nulls are filled** (name -> "Unknown", quantity -> 1, discount -> 0)
   because a missing discount usually means "no discount", not "unknown".
8. **Validate ranges**: sales <= 0, quantity <= 0, discount outside [0,1],
   ship date before order date. These are dropped with counts, never fixed
   by guessing.
9. **Derived columns**: `Year`, `Month`, `Month_Start`, `Year_Month`,
   `Shipping_Days`, `Profit_Margin`, `Flag_Negative_Profit`,
   `Flag_High_Discount`. Flags keep suspicious rows in the data (a loss is a
   business signal) but make them easy to filter.
10. **Sort chronologically** so the file reads like a ledger. `Row ID` survives
    for traceability.
11. **Save as utf-8-sig** so Excel on Windows opens accented names correctly.
    Write a `metadata.json` with source, counts and totals for reproducibility.

## Script 03 - the dashboard

- **Aggregates first, writing second.** `build_aggregates()` produces small
  tables (category, sub-category, region, monthly, top products). The writers
  only format and place them. This keeps the numbers in one place.
- **Values, not formulas.** openpyxl never calculates, so a formula cell would
  show as 0 until Excel recalculates. Writing numbers is portable and always
  correct.
- **Cross-sheet charts.** The Summary charts point at the Category and
  Monthly_Trend tables, so there is only one copy of each table.
- **KPI cards** are six single cells (label / value / note) with a fill and
  border. No merged cells, which avoids a classic openpyxl border bug.
- **Conditional formatting**: colour scale on margin, red font on negative
  profit, data bars on sales. The owner sees problems at a glance.
- **Native Excel charts** (line, bar, pie) with `delete=False` on both axes,
  because openpyxl charts sometimes hide an axis otherwise.
- **Verification at the end**: reopen the file from disk, check sheet names,
  count chart XML files and their types, check KPI cells, check PNG sizes.
  A click that returned is not proof the task succeeded.

## How to explain it in an interview

1. "I read the CSV as text first, because pandas' type guessing once turned a
   ZIP code into a number and dropped a leading zero."
2. "I parse dates from explicit formats and abort on anything I cannot parse,
   rather than silently guessing day vs month."
3. "I flag negative profit and high discounts instead of deleting them: a loss
   is information, not an error."
4. "Margin matters more than sales: a $1M line at 1% margin is worse than a
   $100k line at 40%."
5. "AOV divides by unique orders, not by line items, because one order can
   contain several products."
6. "The workbook uses values, not formulas, because openpyxl does not calculate
   and a formula cell would show as 0 until Excel recalculates."
7. "Every chart is a native Excel chart, not an image, so the owner can hover
   and read exact values."
8. "The whole thing runs offline, with free libraries, and the synthetic demo
   is labelled as such in the workbook, the CSV metadata, the PNGs and the README."

## Common errors and fixes

- `ModuleNotFoundError`: install with `pip install -r requirements.txt` inside
  the venv. Script 01 runs without any package.
- `UnicodeDecodeError`: the file is probably zipped or in another encoding;
  unzip it, or add an encoding to the fallback list.
- `PermissionError` on the xlsx: close it in Excel and rerun.
- Charts with no axes: set `chart.x_axis.delete = False` and the same for y.
- Merged cells with no border: style every cell in the range before merging.
- A chart referencing another sheet works, but the sheet must exist first.

## Exercises (try these to own the code)

1. Change `HIGH_DISCOUNT_THRESHOLD` from 0.40 to 0.30. How do the flag count
   and the "What to do next" text change?
2. Add a 7th KPI card: number of unique customers.
3. Add a new chart: profit by ship mode. Which sheet should it live on?
4. Make the synthetic generator produce 2,000 rows and re-run. Does the
   dashboard still fit on one screen?
5. Replace the teal palette with your own three hex colours in one place.

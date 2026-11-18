# HOW_TO_REBUILD.md - rebuild this project without AI

This guide rebuilds the project by hand. It is written for a beginner who can
read Python but has never built a pipeline. Time estimate: one focused weekend.

Rule zero: **do not copy the final code.** Build the smallest version that runs,
check the output, then add one feature. That is how you learn what each line does.

## Milestone 0 - setup (15 min)

1. Create the folder: `C:\Users\Moaz\freelance-portfolio\proj1` with
   `src`, `data\raw`, `data\clean`, `output`, `docs`.
2. `python -m venv .venv` then `.\.venv\Scripts\Activate.ps1`.
3. `pip install pandas==3.1.1 numpy==2.5.5 openpyxl==3.1.5 matplotlib==3.11.2`.
4. `git init -b main` and create `.gitignore` (venv, raw CSVs, pycache, backups).
5. Get the data: download from Kaggle or run your own generator later. For now,
   save any small CSV as `data\raw\superstore.csv` and inspect it with
   `python -c "import pandas as pd; df=pd.read_csv('data/raw/superstore.csv'); print(df.shape); print(df.dtypes); print(df.isna().sum())"`.

Checkpoint: you can print the shape and columns. If not, fix the path before
moving on.

## Milestone 1 - inspect before cleaning (30 min)

Write a throwaway script that prints: shape, dtypes, null counts, duplicate
count, min/max dates, unique counts for Category/Region. Look at the raw values.
Ask: what is messy? What is already clean? Write down 3 decisions (e.g. "keep
negative profit, flag it").

Checkpoint: you have a written list of cleaning decisions. Do not code yet.

## Milestone 2 - clean (2-3 hours)

Build `src/02_clean.py` in this order, running after each block:

1. Read the CSV with `dtype=str` and an encoding fallback. Print shape.
2. Normalise headers (strip, collapse spaces). Print them.
3. Normalise strings; drop exact duplicates; print how many.
4. Parse dates from one explicit format. Print min/max. If any fail, stop.
5. Convert numerics with `errors="coerce"`; print bad counts.
6. Drop critical nulls; fill optional nulls; print counts.
7. Validate ranges; print dropped counts.
8. Add derived columns; print the new column list.
9. Sort and save as utf-8-sig. Print totals.

Checkpoint after each block: the script runs and prints a count. If a count is
wrong, stop and fix before adding the next block. Never add two blocks at once.

## Milestone 3 - aggregate (1 hour)

Before touching Excel, compute the numbers in a notebook or a `python -c` and
print them: total sales/profit/margin/orders, category table, region table,
monthly table, top 10 products, loss-makers. Check them against the raw data by
hand for one category. If the numbers look wrong, fix step 2; do not patch
numbers in step 3.

## Milestone 4 - Excel skeleton (1 hour)

Create the workbook with openpyxl: one sheet, a title, a small table with
borders, number formats, column widths, freeze panes. Save and open it in Excel
to confirm it opens without a repair warning. This catches 90% of formatting
bugs early.

Key openpyxl facts:
- `ws.merge_cells("A1:D1")` then set `ws["A1"].value`.
- Style every cell in a merged range before merging, or borders disappear.
- `cell.number_format = '"$"#,##0'` for currency, `'0.0%'` for percent.
- Merged cells do not auto-fit row height; set `ws.row_dimensions[r].height`.

## Milestone 5 - charts (1-2 hours)

Add one chart at a time: line (monthly sales), bar (category), pie (region).
Each time: save, open in Excel, confirm it renders. Common fixes:
- Set `chart.x_axis.delete = False` and `chart.y_axis.delete = False`.
- Give the chart a `Reference` with exact `min_row`/`max_row`/`min_col`.
- Put the chart on the sheet that owns its data, or reference the other sheet
  explicitly.
- For coloured bars, use `openpyxl.chart.marker.DataPoint` (note: it lives in
  `openpyxl.chart.marker`, not `bar_chart`).

## Milestone 6 - KPI cards and formatting (1-2 hours)

Add the Summary sheet: 6 KPI cards, "how to read", 3 actions, 3 insights,
conditional formatting, data bars, print setup. Keep cards as single cells (no
merges) to avoid border bugs. Verify by opening in Excel.

## Milestone 7 - PNG charts (1 hour)

matplotlib with `Agg` backend, 3 charts, save at 160 dpi. Watch for:
- Hex colours need a `#` prefix in matplotlib but not in openpyxl.
- `plt.close(fig)` to free memory.
- `bbox_inches="tight"` to keep labels.

## Milestone 8 - docs and verification (1 hour)

Write README (problem, source, tools, run, findings, limits), LEARNING (concepts
+ interview answers), and a verification pass: reopen the xlsx, count charts,
check KPI cells, check PNG sizes. Run the whole pipeline twice (clean output
folder first) to prove it is reproducible. Commit.

## Debugging table

| Symptom | Likely cause | Fix |
|---|---|---|
| `ModuleNotFoundError` | Not in venv | Activate venv or use `.venv\Scripts\python.exe` |
| Excel "repair" warning | Invalid chart ref or merged-cell style | Check refs; style before merge |
| Chart shows no data | Wrong `min_row`/`max_row` | Print the rows; fix the range |
| Chart missing axis | openpyxl default | `chart.x_axis.delete = False` |
| Dates become 1899 numbers | Saved as Excel serial | Format column as date |
| Accented names garbled | No BOM | Save as `utf-8-sig` |
| Totals off by cents | Rounded each line | Keep source precision; round only for display |
| Different row count | Duplicates or dropped rows | Print counts at every step |

## Definition of done

- [ ] `pip install -r requirements.txt` succeeds in a fresh venv
- [ ] `python src/01_check_data.py` with empty `data/raw` exits 1 with instructions
- [ ] `python src/00_generate_sample.py` then `01` `02` `03` all exit 0
- [ ] With the real CSV, all three exit 0 and totals match the published figures
- [ ] `output/Sales_Dashboard.xlsx` opens in Excel without a repair warning
- [ ] 3 PNGs exist and are non-empty
- [ ] `git status` is clean after one commit

## Using this with a client's export (no AI)

Real exports rarely match the Superstore schema. The smallest change is to add
an alias map at the top of `src/02_clean.py`:

```python
COLUMN_ALIASES = {
    "Revenue": "Sales",
    "OrderDate": "Order Date",
    "Qty": "Quantity",
}
```

and rename columns before validation. Then rerun 02 and 03. Change `CURRENCY`
and the threshold constants for that client. Everything else stays the same.

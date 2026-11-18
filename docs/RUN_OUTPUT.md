# Run output (acceptance transcripts)

Both passes were run from the project root with the venv activated.
See `docs/run_synthetic.txt` and `docs/run_real.txt` for the raw console output.

## Pass 0 - empty `data/raw` (the guard)

Command: `python src/01_check_data.py` with an empty `data/raw`.

```
[01] Checking input data...
     Project root: C:\Users\Moaz\freelance-portfolio\proj1
------------------------------------------------------------------------
INPUT DATA NOT FOUND - download instructions
------------------------------------------------------------------------
Expected folder : C:\Users\Moaz\freelance-portfolio\proj1\data\raw
Expected file   : superstore.csv  (the Kaggle 'Sample - Superstore.csv')
...
No Kaggle account? Generate a synthetic demo instead (clearly labelled):
  python src/00_generate_sample.py
------------------------------------------------------------------------
EXIT CODE: 1
```

This exit code 1 is the guard working as designed, not a failure.

## Pass 1 - synthetic data

```
$ python src/00_generate_sample.py
SYNTHETIC DATA GENERATED - seed=42 - NOT real business data
Rows written : 1492
...
$ python src/01_check_data.py
[01] SYNTHETIC dataset (synthetic_superstore.csv): ...
     OK - 1,492 data rows, 19 columns, 0.28 MB
EXIT: 0
$ python src/02_clean.py
     duplicates removed: 10  (rows now 1,482)
     rows in/out   : 1,492 -> 1,478
     total sales   : 1,587,468.00  (synthetic)
     total profit  : 242,200.00
     margin        : 15.26%
     orders        : 599
EXIT: 0
$ python src/03_report.py
     sheets OK: ['Summary', 'Category', 'Region', 'Monthly_Trend', 'Top_Products']
     chart XML files: 6  -> {'barChart': 3, 'lineChart': 2, 'pieChart': 1}
     KPIs OK: sales=1,587,468 profit=242,200 margin=15.26%
     PNG OK: 01_monthly_sales_profit.png (1842x963, 171 KB)
     PNG OK: 02_subcategory_profit.png (1736x1118, 86 KB)
     PNG OK: 03_region_share.png (1439x1040, 84 KB)
     conditional-format rules: 3
     verification passed
EXIT: 0
```

## Pass 2 - real Kaggle Superstore CSV

```
$ python src/01_check_data.py
[01] REAL dataset (superstore.csv): ...
     OK - 9,994 data rows, 21 columns, 2.18 MB
EXIT: 0
$ python src/02_clean.py
     duplicates removed: 0
     rows in/out   : 9,994 -> 9,994
     total sales   : 2,297,200.86
     total profit  : 286,397.02
     margin        : 12.47%
     orders        : 5,009
     loss lines    : 1,871
     high-discount : 1,139
EXIT: 0
$ python src/03_report.py
     sheets OK: ['Summary', 'Category', 'Region', 'Monthly_Trend', 'Top_Products']
     chart XML files: 6  -> {'barChart': 3, 'lineChart': 2, 'pieChart': 1}
     KPIs OK: sales=2,297,201 profit=286,397 margin=12.47%
     PNG OK: 01_monthly_sales_profit.png (1842x963, 171 KB)
     PNG OK: 02_subcategory_profit.png (1736x1118, 86 KB)
     PNG OK: 03_region_share.png (1439x1040, 84 KB)
     conditional-format rules: 3
     verification passed
EXIT: 0
```

Totals match the published Superstore figures ($2,297,200.86 sales / $286,397.02 profit).

## Environment

- Python 3.15.2, Windows 11, pip 26.3
- pandas 3.1.1, numpy 2.5.5, openpyxl 3.1.5, matplotlib 3.11.2
- `pip check`: no broken requirements

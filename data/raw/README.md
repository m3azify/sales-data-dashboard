# Raw data folder

This folder holds the input CSVs. They are git-ignored (size + licensing).

- **Real data:** download `Sample - Superstore.csv` from
  https://www.kaggle.com/datasets/vivek468/superstore-dataset-final
  and save it here as `superstore.csv`.
- **Demo data:** run `python src/00_generate_sample.py` to create
  `synthetic_superstore.csv` (clearly labelled synthetic everywhere).

If both files exist, the pipeline uses the real one. Run `python src/01_check_data.py`
to confirm which one will be used.

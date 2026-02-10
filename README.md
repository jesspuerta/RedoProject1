# RedoProject1

Deterministic GitHub-based analysis pipeline for ranking programs/courses from the **2024 Graduate Program Exit Survey** dataset.

## Repository structure

- `data/` - input dataset (`Grad Program Exit Survey Data 2024.xlsx`)
- `src/` - analysis code
- `outputs/` - generated rank outputs
- `.github/workflows/` - CI workflow for PR-triggered execution

## Local run

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Run the ranking pipeline:

   ```bash
   python src/rank_courses_2024.py
   ```

3. Generated files:
   - `outputs/2024_rank_order.csv`
   - `outputs/2024_rank_order.md`
   - `outputs/2024_rank_order.png`

## Determinism and ranking rules

The script is deterministic and does not use randomness.

- It reads the Excel file with the `openpyxl` engine.
- It prioritizes 2024 sheet detection, then `Year == 2024` row filtering when available.
- It auto-detects rating/preference columns using numeric/Likert heuristics and excludes likely metadata/free-text fields.
- It standardizes Likert strings into numeric scores and safely coerces numeric-like values.
- It ranks by:
  1. Higher `mean_score`
  2. Higher `n_responses`
  3. Alphabetical `item`

## GitHub Actions workflow

Workflow: `.github/workflows/rank_courses_2024.yml`

On every pull request (and manual `workflow_dispatch`), it:

1. Checks out the repository
2. Sets up Python 3.11
3. Installs `requirements.txt`
4. Runs `python src/rank_courses_2024.py`
5. Verifies required outputs exist
6. Uploads `outputs/` as artifact: `rank_order_outputs_2024`

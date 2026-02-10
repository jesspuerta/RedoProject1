from __future__ import annotations

import logging
import math
import re
from pathlib import Path
from textwrap import fill

import matplotlib.pyplot as plt
import pandas as pd

DATA_PATH = Path("data/Grad Program Exit Survey Data 2024.xlsx")
OUTPUT_DIR = Path("outputs")
CSV_OUTPUT = OUTPUT_DIR / "2024_rank_order.csv"
MD_OUTPUT = OUTPUT_DIR / "2024_rank_order.md"
PNG_OUTPUT = OUTPUT_DIR / "2024_rank_order.png"

LIKERT_MAPPING = {
    # Agreement scale (high is positive)
    "strongly agree": 5,
    "agree": 4,
    "neutral": 3,
    "neither agree nor disagree": 3,
    "disagree": 2,
    "strongly disagree": 1,
    # Quality scale
    "excellent": 5,
    "very good": 4,
    "good": 3,
    "fair": 2,
    "poor": 1,
    # Satisfaction scale
    "very satisfied": 5,
    "satisfied": 4,
    "somewhat satisfied": 3,
    "somewhat dissatisfied": 2,
    "dissatisfied": 2,
    "very dissatisfied": 1,
    # Frequency/extent scale
    "always": 5,
    "often": 4,
    "sometimes": 3,
    "rarely": 2,
    "never": 1,
    "to a great extent": 5,
    "to a moderate extent": 4,
    "to some extent": 3,
    "to a little extent": 2,
    "not at all": 1,
}

LIKERT_KEYWORDS = set(LIKERT_MAPPING.keys())
EXCLUDE_NAME_HINTS = (
    "timestamp",
    "time stamp",
    "id",
    "email",
    "name",
    "comment",
    "feedback",
    "suggest",
    "other",
    "why",
    "explain",
    "text",
)


def normalize_string(value: object) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def choose_sheet(path: Path) -> str | int:
    excel_file = pd.ExcelFile(path, engine="openpyxl")
    year_sheets = [s for s in excel_file.sheet_names if "2024" in s]
    if year_sheets:
        selected = sorted(year_sheets)[0]
        logging.info("Using 2024-specific sheet: %s", selected)
        return selected

    selected = excel_file.sheet_names[0]
    logging.info("No 2024-specific sheet found. Using first sheet: %s", selected)
    return selected


def filter_to_2024(df: pd.DataFrame) -> pd.DataFrame:
    year_candidates = [c for c in df.columns if "year" in str(c).strip().lower()]
    if not year_candidates:
        logging.info("No year column found; assuming dataset is already 2024-only.")
        return df

    year_col = year_candidates[0]
    year_series = pd.to_numeric(df[year_col], errors="coerce")
    filtered = df[year_series == 2024]
    if filtered.empty:
        logging.info(
            "Year column '%s' exists but no rows equal 2024; proceeding with unfiltered data.",
            year_col,
        )
        return df

    logging.info("Filtered to Year == 2024 using column '%s'. Rows: %d", year_col, len(filtered))
    return filtered


def score_series(series: pd.Series) -> pd.Series:
    normalized = series.map(normalize_string)
    mapped = normalized.map(LIKERT_MAPPING)

    numeric = pd.to_numeric(series, errors="coerce")
    scored = mapped.where(mapped.notna(), numeric)

    return scored


def is_metadata_column(name: str, series: pd.Series) -> bool:
    lower_name = name.strip().lower()
    if any(hint in lower_name for hint in EXCLUDE_NAME_HINTS):
        return True

    non_null = series.dropna()
    if non_null.empty:
        return True

    normalized = non_null.map(normalize_string)
    avg_len = normalized.str.len().mean()
    unique_ratio = normalized.nunique(dropna=True) / len(normalized)

    # Heuristic for free-text fields
    if avg_len > 55 and unique_ratio > 0.75:
        return True

    return False


def detect_rating_columns(df: pd.DataFrame) -> list[str]:
    rating_columns: list[str] = []

    for col in df.columns:
        name = str(col)
        series = df[col]

        if is_metadata_column(name, series):
            continue

        non_null = series.dropna()
        if non_null.empty:
            continue

        normalized = non_null.map(normalize_string)
        numeric_ratio = pd.to_numeric(non_null, errors="coerce").notna().mean()
        likert_ratio = normalized.isin(LIKERT_KEYWORDS).mean()

        if numeric_ratio >= 0.7 or likert_ratio >= 0.4:
            rating_columns.append(name)

    return sorted(rating_columns)


def build_ranking(df: pd.DataFrame, rating_columns: list[str]) -> pd.DataFrame:
    rows = []
    for col in rating_columns:
        scored = score_series(df[col]).dropna()
        if scored.empty:
            continue

        rows.append(
            {
                "item": str(col),
                "mean_score": round(float(scored.mean()), 4),
                "n_responses": int(scored.count()),
            }
        )

    ranking = pd.DataFrame(rows)
    if ranking.empty:
        return ranking

    ranking = ranking.sort_values(
        by=["mean_score", "n_responses", "item"],
        ascending=[False, False, True],
        kind="mergesort",
    ).reset_index(drop=True)
    ranking["rank"] = ranking.index + 1

    return ranking[["item", "mean_score", "n_responses", "rank"]]


def write_markdown_summary(ranking: pd.DataFrame, detected_cols: list[str]) -> None:
    top_n = min(10, len(ranking))
    preview = ranking.head(top_n)

    lines = [
        "# 2024 Program/Course Rank Order",
        "",
        "This ranking is based on average standardized ratings from the 2024 exit survey data.",
        "Tie breaks are resolved deterministically by higher response count, then alphabetical item name.",
        "",
        "## Detected rating/preference columns",
        "",
        *[f"- {col}" for col in detected_cols],
        "",
        f"## Top {top_n} items",
        "",
        "| rank | item | mean_score | n_responses |",
        "|---:|---|---:|---:|",
    ]

    for row in preview.itertuples(index=False):
        lines.append(f"| {row.rank} | {row.item} | {row.mean_score:.4f} | {row.n_responses} |")

    MD_OUTPUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def wrap_labels(labels: list[str], width: int = 35) -> list[str]:
    return [fill(label, width=width) for label in labels]


def build_chart(ranking: pd.DataFrame) -> None:
    top_n = min(15, len(ranking))
    top = ranking.head(top_n).copy()
    top = top.sort_values(by="mean_score", ascending=True)

    labels = wrap_labels(top["item"].tolist(), width=35)
    scores = top["mean_score"].tolist()

    plt.figure(figsize=(12, max(6, top_n * 0.5)))
    bars = plt.barh(labels, scores, color="#2E86AB")
    plt.xlabel("Average Rating")
    plt.ylabel("Program/Course Item")
    plt.title("2024 Exit Survey: Program/Course Ranking by Average Rating")

    for bar, score in zip(bars, scores):
        plt.text(
            bar.get_width() + 0.02,
            bar.get_y() + bar.get_height() / 2,
            f"{score:.2f}",
            va="center",
            fontsize=9,
        )

    plt.xlim(0, max(scores) + 0.5)
    plt.tight_layout()
    plt.savefig(PNG_OUTPUT, dpi=200)
    plt.close()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Dataset not found at {DATA_PATH}")

    selected_sheet = choose_sheet(DATA_PATH)
    df = pd.read_excel(DATA_PATH, sheet_name=selected_sheet, engine="openpyxl")
    df_2024 = filter_to_2024(df)

    rating_columns = detect_rating_columns(df_2024)
    if not rating_columns:
        raise ValueError("No rating/preference columns were detected.")

    logging.info("Detected rating/preference columns (%d): %s", len(rating_columns), rating_columns)
    logging.info("Likert mapping used: %s", LIKERT_MAPPING)

    ranking = build_ranking(df_2024, rating_columns)
    if ranking.empty:
        raise ValueError("No rankable data after standardization.")

    ranking.to_csv(CSV_OUTPUT, index=False)
    write_markdown_summary(ranking, rating_columns)
    build_chart(ranking)

    logging.info("Saved ranking CSV to %s", CSV_OUTPUT)
    logging.info("Saved ranking markdown to %s", MD_OUTPUT)
    logging.info("Saved ranking chart to %s", PNG_OUTPUT)


if __name__ == "__main__":
    main()

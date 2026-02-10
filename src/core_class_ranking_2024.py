from __future__ import annotations

import logging
from pathlib import Path
from textwrap import fill

import matplotlib.pyplot as plt
import pandas as pd

DATA_PATH = Path("data/Grad Program Exit Survey Data 2024.xlsx")
OUTPUT_DIR = Path("outputs")
CSV_OUTPUT = OUTPUT_DIR / "core_class_ranking_2024.csv"
MD_OUTPUT = OUTPUT_DIR / "core_class_ranking_2024.md"
PNG_OUTPUT = OUTPUT_DIR / "core_class_ranking_2024.png"

TIME_HEADER_HINTS = (
    "duration",
    "time",
    "timestamp",
    "seconds",
    "minutes",
    "hours",
)


def load_core_columns(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl", usecols="L:S")
    logging.info("Loaded columns L:S only. Detected headers: %s", list(df.columns))
    return df


def detect_duration_like_columns(df: pd.DataFrame) -> list[str]:
    flagged: list[str] = []

    for col in df.columns:
        series = df[col].dropna()
        if series.empty:
            continue

        header_lower = str(col).strip().lower()
        if any(hint in header_lower for hint in TIME_HEADER_HINTS):
            flagged.append(str(col))
            logging.warning("Excluding '%s' because header looks time/duration-like.", col)
            continue

        as_num = pd.to_numeric(series, errors="coerce")
        numeric_ratio = as_num.notna().mean()
        if numeric_ratio >= 0.8:
            median_val = float(as_num.median())
            max_val = float(as_num.max())
            if median_val > 100 and max_val > 1000:
                flagged.append(str(col))
                logging.warning(
                    "Excluding '%s' because values look like duration (median=%.2f, max=%.2f).",
                    col,
                    median_val,
                    max_val,
                )
                continue

        as_dt = pd.to_datetime(series, errors="coerce")
        datetime_ratio = as_dt.notna().mean()
        if datetime_ratio >= 0.6:
            flagged.append(str(col))
            logging.warning(
                "Excluding '%s' because values parse as timestamps (ratio=%.2f).",
                col,
                datetime_ratio,
            )

    return sorted(set(flagged))


def sanitize_ranks(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int], int]:
    k_classes = len(df.columns)
    numeric = df.apply(pd.to_numeric, errors="coerce")

    out_of_range_counts: dict[str, int] = {}
    for col in numeric.columns:
        series = numeric[col]
        out_of_range = series.notna() & ~series.between(1, k_classes)
        out_of_range_counts[str(col)] = int(out_of_range.sum())

        non_integer = series.notna() & (series % 1 != 0)
        invalid_mask = out_of_range | non_integer
        numeric.loc[invalid_mask, col] = pd.NA

    return numeric, out_of_range_counts, k_classes


def log_validation_checks(ranks: pd.DataFrame, out_of_range_counts: dict[str, int], k_classes: int) -> None:
    rows = len(ranks)
    if rows == 0:
        logging.warning("No rows found in L:S.")
        return

    has_any_rank_pct = float(ranks.notna().any(axis=1).mean() * 100)

    unique_rows = 0
    rows_with_values = 0
    for _, row in ranks.iterrows():
        vals = row.dropna().astype(int).tolist()
        if vals:
            rows_with_values += 1
            if len(vals) == len(set(vals)):
                unique_rows += 1

    unique_pct = (unique_rows / rows_with_values * 100) if rows_with_values else 0.0

    total_out_of_range = sum(out_of_range_counts.values())

    logging.info("Validation: K classes=%d", k_classes)
    logging.info("Validation: %.2f%% of rows have at least one rank value.", has_any_rank_pct)
    logging.info(
        "Validation: %.2f%% of rows with rank values have unique ranks across classes.",
        unique_pct,
    )
    logging.info("Validation: out-of-range rank entries removed=%d", total_out_of_range)
    if total_out_of_range:
        logging.info("Validation detail (out-of-range by class): %s", out_of_range_counts)


def compute_ranking(ranks: pd.DataFrame) -> pd.DataFrame:
    k_classes = len(ranks.columns)
    rows: list[dict[str, float | int | str]] = []

    for col in ranks.columns:
        series = ranks[col].dropna().astype(int)
        n = int(series.count())

        if n == 0:
            mean_rank = float("nan")
            median_rank = float("nan")
            pct_top1 = 0.0
            pct_top3 = 0.0
        else:
            mean_rank = float(series.mean())
            median_rank = float(series.median())
            pct_top1 = float((series == 1).mean() * 100)
            top_n = min(3, k_classes)
            pct_top3 = float((series <= top_n).mean() * 100)

        rows.append(
            {
                "class": str(col),
                "mean_rank": mean_rank,
                "median_rank": median_rank,
                "n_responses": n,
                "pct_top1": pct_top1,
                "pct_top3": pct_top3,
            }
        )

    ranking = pd.DataFrame(rows)
    ranking = ranking.sort_values(
        by=["mean_rank", "median_rank", "n_responses", "class"],
        ascending=[True, True, False, True],
        kind="mergesort",
        na_position="last",
    ).reset_index(drop=True)
    ranking["rank"] = ranking.index + 1

    return ranking[
        ["class", "mean_rank", "median_rank", "n_responses", "pct_top1", "pct_top3", "rank"]
    ]


def build_markdown_summary(ranking: pd.DataFrame, excluded_columns: list[str]) -> str:
    lines = [
        "# 2024 Exit Survey: CORE Class Ranking (1st → Last)",
        "",
        "Ranking computed from Excel columns L:S only, interpreted as rank-style responses (1=best).",
        "",
    ]

    if excluded_columns:
        lines.extend(
            [
                "## Excluded columns",
                "",
                "The following L:S columns were excluded because they appeared time/duration-like:",
                *[f"- {col}" for col in excluded_columns],
                "",
            ]
        )

    lines.extend(
        [
            "## Ranking table",
            "",
            "| rank | class | mean_rank | median_rank | n_responses | pct_top1 | pct_top3 |",
            "|---:|---|---:|---:|---:|---:|---:|",
        ]
    )

    for row in ranking.itertuples(index=False):
        row_data = row._asdict()
        lines.append(
            f"| {row_data['rank']} | {row_data['class']} | {row_data['mean_rank']:.3f} | {row_data['median_rank']:.3f} | "
            f"{row_data['n_responses']} | {row_data['pct_top1']:.1f}% | {row_data['pct_top3']:.1f}% |"
        )

    top3 = ranking.head(min(3, len(ranking)))["class"].tolist()
    bottom3 = ranking.tail(min(3, len(ranking)))["class"].tolist()
    gap = (
        float(ranking["mean_rank"].max() - ranking["mean_rank"].min())
        if not ranking.empty
        else float("nan")
    )

    lines.extend(
        [
            "",
            "## Insights",
            "",
            f"- Top classes (best mean rank): {', '.join(top3) if top3 else 'N/A'}.",
            f"- Bottom classes (worst mean rank): {', '.join(bottom3) if bottom3 else 'N/A'}.",
            f"- Mean-rank spread from best to worst class: {gap:.3f}.",
            (
                f"- Response counts range from {int(ranking['n_responses'].min())} to "
                f"{int(ranking['n_responses'].max())}."
                if not ranking.empty
                else "- Response counts unavailable."
            ),
        ]
    )

    return "\n".join(lines) + "\n"


def build_plot(ranking: pd.DataFrame) -> None:
    ordered = ranking.sort_values("rank", ascending=True).copy()
    y_positions = list(range(len(ordered)))

    labels = [f"#{int(r)} {fill(c, width=34)}" for r, c in zip(ordered["rank"], ordered["class"])]
    values = ordered["mean_rank"].tolist()

    plt.figure(figsize=(12, max(6, len(ordered) * 0.7)))
    for y, x in zip(y_positions, values):
        plt.hlines(y=y, xmin=0, xmax=x, color="#8ab6d6", linewidth=2)
    plt.scatter(values, y_positions, color="#1f4e79", s=70, zorder=3)

    for y, x, r in zip(y_positions, values, ordered["rank"]):
        plt.text(x + 0.03, y, f"{int(r)}", va="center", fontsize=10)

    plt.yticks(y_positions, labels)
    plt.gca().invert_yaxis()
    plt.xlabel("mean_rank (lower is better)")
    plt.title("2024 Exit Survey: CORE Class Ranking (1st → Last)")
    plt.grid(axis="x", linestyle="--", alpha=0.3)
    plt.tight_layout()
    plt.savefig(PNG_OUTPUT, dpi=220)
    plt.close()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Dataset not found: {DATA_PATH}")

    df_ls = load_core_columns(DATA_PATH)
    excluded = detect_duration_like_columns(df_ls)
    df_core = df_ls.drop(columns=excluded, errors="ignore")

    if df_core.empty:
        raise ValueError("No usable CORE class columns left in L:S after exclusions.")

    ranks, out_of_range_counts, k_classes = sanitize_ranks(df_core)
    log_validation_checks(ranks, out_of_range_counts, k_classes)

    ranking = compute_ranking(ranks)
    ranking.to_csv(CSV_OUTPUT, index=False)

    md = build_markdown_summary(ranking, excluded)
    MD_OUTPUT.write_text(md, encoding="utf-8")

    build_plot(ranking)

    logging.info("Saved: %s", CSV_OUTPUT)
    logging.info("Saved: %s", MD_OUTPUT)
    logging.info("Saved: %s", PNG_OUTPUT)


if __name__ == "__main__":
    main()

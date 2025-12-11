"""Perform random-effects meta-analysis of TB delay metrics."""
from __future__ import annotations

import logging
import math
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = PROJECT_ROOT / "data" / "processed" / "lit_delay_extracted.csv"
OUTPUT_TABLE = PROJECT_ROOT / "data" / "processed" / "meta_delay_results.csv"
FIGURES_DIR = PROJECT_ROOT / "output" / "figures"
LOG_PATH = PROJECT_ROOT / "data" / "processed" / "meta_analysis.log"

DELAY_TYPES = {
    "patient_delay_days": "Patient delay (days)",
    "diagnostic_delay_days": "Diagnostic delay (days)",
    "treatment_delay_days": "Treatment delay (days)",
    "total_delay_days": "Total delay (days)",
}


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("meta_analysis")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(formatter)
    file_handler = logging.FileHandler(LOG_PATH)
    file_handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.addHandler(file_handler)
    return logger


def load_data() -> pd.DataFrame:
    if not DATA_PATH.exists() or DATA_PATH.stat().st_size == 0:
        return pd.DataFrame()
    return pd.read_csv(DATA_PATH)


def compute_standard_error(row: pd.Series, column: str) -> float:
    se_column = column.replace("_days", "_se")
    if se_column in row and not pd.isna(row[se_column]):
        return float(row[se_column])
    sd_column = column.replace("_days", "_sd")
    if sd_column in row and "sample_size" in row and row.get("sample_size"):
        try:
            return float(row[sd_column]) / math.sqrt(float(row["sample_size"]))
        except (ValueError, ZeroDivisionError):
            return math.nan
    return math.nan


def dersimonian_laird(effects: np.ndarray, variances: np.ndarray) -> Dict[str, float]:
    weights = 1 / variances
    fixed_effect = np.sum(weights * effects) / np.sum(weights)
    q = np.sum(weights * (effects - fixed_effect) ** 2)
    df = len(effects) - 1
    c = np.sum(weights) - (np.sum(weights ** 2) / np.sum(weights))
    tau_sq = max(0.0, (q - df) / c) if c > 0 else 0.0
    weights_random = 1 / (variances + tau_sq)
    pooled = np.sum(weights_random * effects) / np.sum(weights_random)
    pooled_se = math.sqrt(1 / np.sum(weights_random))
    ci_low = pooled - 1.96 * pooled_se
    ci_high = pooled + 1.96 * pooled_se
    return {
        "effect": pooled,
        "se": pooled_se,
        "ci_low": ci_low,
        "ci_high": ci_high,
        "tau_sq": tau_sq,
        "q": q,
        "k": len(effects),
    }


def run_meta_analysis(df: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    results = []
    for column, label in DELAY_TYPES.items():
        if column not in df.columns:
            logger.warning("Column %s missing; skipping.", column)
            continue
        cols = [
            column,
            *[c for c in ["pmid", "state", "study_year", "setting_public_private", "sample_size"] if c in df.columns],
        ]
        subset = df[cols].copy()
        for aux_col in [
            column.replace("_days", "_se"),
            column.replace("_days", "_sd"),
        ]:
            if aux_col in df.columns and aux_col not in subset.columns:
                subset[aux_col] = df.loc[subset.index, aux_col]
        subset = subset.dropna(subset=[column])
        if subset.empty:
            logger.warning("No data for %s", column)
            continue
        se_values = subset.apply(lambda row: compute_standard_error(row, column), axis=1)
        subset["variance"] = se_values ** 2
        subset.loc[subset["variance"].isna() | (subset["variance"] == 0), "variance"] = 1.0
        effects = subset[column].astype(float).values
        variances = subset["variance"].astype(float).values
        stats = dersimonian_laird(effects, variances)
        stats.update({"delay_type": label})
        results.append(stats)
        create_forest_plot(subset, label, stats, logger)
    return pd.DataFrame(results)


def create_forest_plot(subset: pd.DataFrame, label: str, stats: Dict[str, float], logger: logging.Logger) -> None:
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(8, max(4, 0.3 * len(subset))))
    y_positions = range(len(subset))
    effects = subset.iloc[:, 0].astype(float)
    se_values = subset["variance"].apply(math.sqrt)
    ax.errorbar(effects, y_positions, xerr=1.96 * se_values, fmt="o", color="steelblue")
    ax.axvline(stats["effect"], color="darkred", linestyle="--", label="Pooled")
    labels = subset["pmid"].astype(str) if "pmid" in subset.columns else subset.index.astype(str)
    ax.set_yticks(list(y_positions))
    ax.set_yticklabels(labels)
    ax.set_xlabel(label)
    ax.set_title(f"Forest plot: {label}")
    ax.legend()
    fig.tight_layout()
    output_file = FIGURES_DIR / f"forest_{label.replace(' ', '_').lower()}.png"
    fig.savefig(output_file, dpi=300)
    plt.close(fig)
    logger.info("Saved forest plot to %s", output_file)


def main() -> None:
    logger = configure_logging()
    df = load_data()
    if df.empty:
        logger.error("Literature extraction file missing or empty; cannot run meta-analysis.")
        OUTPUT_TABLE.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(columns=["delay_type", "effect", "ci_low", "ci_high", "tau_sq", "k"]).to_csv(
            OUTPUT_TABLE, index=False
        )
        return
    results = run_meta_analysis(df, logger)
    OUTPUT_TABLE.parent.mkdir(parents=True, exist_ok=True)
    results.to_csv(OUTPUT_TABLE, index=False)
    logger.info("Meta-analysis results saved to %s", OUTPUT_TABLE)


if __name__ == "__main__":
    main()

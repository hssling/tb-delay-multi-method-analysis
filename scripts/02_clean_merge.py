"""Clean and merge processed datasets into a state-year panel."""
from __future__ import annotations

import logging
import sys
from functools import reduce
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
OUTPUT_PATH = PROCESSED_DIR / "state_year_panel.csv"
LOG_PATH = PROCESSED_DIR / "clean_merge.log"

PROCESSED_FILES = {
    "who_tb_global": PROCESSED_DIR / "who_tb_global_clean.csv",
    "india_tb_reports": PROCESSED_DIR / "india_tb_reports_clean.csv",
    "india_tb_notifications_2025": PROCESSED_DIR / "india_tb_notifications_2025_clean.csv",
    "tb_prevalence": PROCESSED_DIR / "tb_prevalence_clean.csv",
    "nfhs": PROCESSED_DIR / "nfhs_clean.csv",
    "census": PROCESSED_DIR / "census_clean.csv",
}

STATE_ALIASES = {
    "andaman and nicobar islands": "andaman & nicobar islands",
    "dadra and nagar haveli and daman and diu": "dadra & nagar haveli and daman & diu",
    "odisha": "odisha",
    "orissa": "odisha",
    "nct of delhi": "delhi",
    "delhi": "delhi",
    "uttaranchal": "uttarakhand",
    "pondicherry": "puducherry",
    "telangana": "telangana",
}

STATE_CANDIDATES = {
    "state",
    "state_ut",
    "stateut",
    "statenames",
    "state_name",
    "state/ut",
    "name",
}
YEAR_CANDIDATES = {"year", "fy", "financial_year", "survey_year", "time"}


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("clean_merge")
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


def normalize_state_name(value: str) -> str:
    if value is None:
        return "unknown"
    normalized = str(value).strip().lower()
    normalized = normalized.replace("&", "and")
    return STATE_ALIASES.get(normalized, normalized.title())


def detect_column(df: pd.DataFrame, candidates: set[str]) -> Optional[str]:
    for col in df.columns:
        if col.lower() in candidates:
            return col
    return None


def load_dataset(path: Path, dataset_name: str, logger: logging.Logger) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        logger.warning("Processed file %s missing or empty.", path)
        return pd.DataFrame()
    df = pd.read_csv(path)
    if df.empty:
        logger.warning("Processed file %s contains no data.", path)
    return standardize_dataset(df, dataset_name, logger)


def standardize_dataset(df: pd.DataFrame, dataset_name: str, logger: logging.Logger) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    state_col = detect_column(df, STATE_CANDIDATES)
    if not state_col:
        logger.error("Dataset %s lacks a recognizable state column.", dataset_name)
        return pd.DataFrame()
    df["state"] = df[state_col].apply(normalize_state_name)
    year_col = detect_column(df, YEAR_CANDIDATES)
    if year_col:
        df["year"] = df[year_col]
    else:
        logger.info("Dataset %s lacks year column; merging on state only.", dataset_name)
    drop_cols = {state_col}
    if year_col:
        drop_cols.add(year_col)
    value_cols = [
        c for c in df.columns if c not in drop_cols.union({"state", "year"})
    ]
    ordered_cols = ["state"] + (["year"] if "year" in df.columns else []) + value_cols
    df = df[ordered_cols]
    renamed_cols = {
        col: (col if col in {"state", "year"} else f"{dataset_name}_{col}")
        for col in df.columns
    }
    df.rename(columns=renamed_cols, inplace=True)
    return df


def merge_datasets(datasets: Dict[str, pd.DataFrame], logger: logging.Logger) -> pd.DataFrame:
    frames: List[pd.DataFrame] = [df for df in datasets.values() if not df.empty]
    if not frames:
        logger.warning("No datasets available to merge; writing empty panel.")
        return pd.DataFrame(columns=["state", "year"])

    def merge_two(left: pd.DataFrame, right: pd.DataFrame) -> pd.DataFrame:
        if right.empty:
            return left
        keys = [k for k in ["state", "year"] if k in left.columns and k in right.columns]
        if not keys:
            keys = ["state"]
        logger.info("Merging on keys: %s", keys)
        return pd.merge(left, right, on=keys, how="outer")

    merged = reduce(merge_two, frames)
    merged.sort_values(by=[c for c in ["year", "state"] if c in merged.columns], inplace=True)
    return merged


def main() -> None:
    logger = configure_logging()
    logger.info("Loading processed datasets for merging.")
    datasets = {
        name: load_dataset(path, name, logger)
        for name, path in PROCESSED_FILES.items()
    }
    panel = merge_datasets(datasets, logger)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    panel.to_csv(OUTPUT_PATH, index=False)
    logger.info("State-year panel saved to %s with %s rows.", OUTPUT_PATH, len(panel))


if __name__ == "__main__":
    main()

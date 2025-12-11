"""Ingest public TB datasets and export cleaned CSVs."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Callable, Dict

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = PROJECT_ROOT / "data" / "raw"
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
LOG_PATH = PROCESSED_DIR / "ingest_sources.log"

DATASETS = {
    "who_tb_global": RAW_DIR / "who_tb_global.csv",
    "india_tb_reports": RAW_DIR / "india_tb_reports_statewise.xlsx",
    "india_tb_notifications_2025": RAW_DIR / "india_tb_notifications_2025.csv",
    "tb_prevalence": RAW_DIR / "tb_prevalence_survey.xlsx",
    "nfhs": RAW_DIR / "nfhs_state_indicators.csv",
    "census": RAW_DIR / "census_state_indicators.csv",
}


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("ingest_sources")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    file_handler = logging.FileHandler(LOG_PATH)
    file_handler.setFormatter(formatter)
    logger.addHandler(stream)
    logger.addHandler(file_handler)
    return logger


def read_file(path: Path, loader: Callable[[Path], pd.DataFrame], logger: logging.Logger) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        logger.warning("File %s missing or empty; returning empty DataFrame.", path)
        return pd.DataFrame()
    try:
        df = loader(path)
        if df.empty:
            logger.warning("File %s loaded but contains no rows.", path)
        return df
    except Exception as exc:  # noqa: BLE001
        logger.exception("Failed to load %s: %s", path, exc)
        return pd.DataFrame()


def load_who(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    columns = {c: c.strip().lower() for c in df.columns}
    df.rename(columns=columns, inplace=True)
    if "country" in df.columns:
        df = df[df["country"].str.lower() == "india"]
    return df


def load_excel(path: Path) -> pd.DataFrame:
    return pd.read_excel(path)


def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path)


def clean_generic(df: pd.DataFrame, dataset_name: str) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    df["source_dataset"] = dataset_name
    return df


def save_dataset(df: pd.DataFrame, name: str, logger: logging.Logger) -> None:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    output_path = PROCESSED_DIR / f"{name}_clean.csv"
    df.to_csv(output_path, index=False)
    logger.info("Saved %s rows to %s", len(df), output_path)


def ingest_datasets(selected: Dict[str, Path], logger: logging.Logger) -> None:
    loaders: Dict[str, Callable[[Path], pd.DataFrame]] = {
        "who_tb_global": load_who,
        "india_tb_reports": load_excel,
        "india_tb_notifications_2025": load_csv,
        "tb_prevalence": load_excel,
        "nfhs": load_csv,
        "census": load_csv,
    }
    for name, path in selected.items():
        df = read_file(path, loaders[name], logger)
        df = clean_generic(df, name)
        save_dataset(df, name, logger)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Ingest WHO/NFHS/Census tuberculosis datasets."
    )
    parser.add_argument(
        "--skip-who",
        action="store_true",
        help="Skip WHO TB global dataset processing if already done.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logger = configure_logging()
    logger.info("Starting ingestion of TB data sources.")
    datasets_to_process = DATASETS.copy()
    if args.skip_who and "who_tb_global" in datasets_to_process:
        logger.info("Skip WHO dataset flag detected; removing from processing queue.")
        datasets_to_process.pop("who_tb_global")

    ingest_datasets(datasets_to_process, logger)
    logger.info("Ingestion complete.")


if __name__ == "__main__":
    main()

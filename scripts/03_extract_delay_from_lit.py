"""Create and update literature extraction template for TB delay metrics."""
from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Dict, List

import pandas as pd
import re

PROJECT_ROOT = Path(__file__).resolve().parents[1]
LIT_DB_PATH = PROJECT_ROOT / "lit" / "lit_db.csv"
TEMPLATE_PATH = PROJECT_ROOT / "lit" / "extracted_studies_template.csv"
OUTPUT_PATH = PROJECT_ROOT / "data" / "processed" / "lit_delay_extracted.csv"
LOG_PATH = PROJECT_ROOT / "data" / "processed" / "extract_delay.log"

DELAY_COLUMNS = [
    "patient_delay_days",
    "patient_delay_se",
    "diagnostic_delay_days",
    "diagnostic_delay_se",
    "treatment_delay_days",
    "treatment_delay_se",
    "total_delay_days",
    "total_delay_se",
]

TEMPLATE_COLUMNS = [
    "pmid",
    "title",
    "state",
    "study_year",
    "setting_public_private",
    "sample_size",
    *DELAY_COLUMNS,
    "notes",
    "extraction_status",
]

STATE_VARIANTS: Dict[str, List[str]] = {
    "andhra pradesh": ["andhra pradesh"],
    "arunachal pradesh": ["arunachal"],
    "assam": ["assam"],
    "bihar": ["bihar"],
    "chhattisgarh": ["chhattisgarh"],
    "delhi": ["delhi", "nct of delhi"],
    "goa": ["goa"],
    "gujarat": ["gujarat"],
    "haryana": ["haryana"],
    "himachal pradesh": ["himachal"],
    "jharkhand": ["jharkhand"],
    "karnataka": ["karnataka"],
    "kerala": ["kerala"],
    "madhya pradesh": ["madhya pradesh", "mp "],
    "maharashtra": ["maharashtra"],
    "manipur": ["manipur"],
    "meghalaya": ["meghalaya"],
    "mizoram": ["mizoram"],
    "nagaland": ["nagaland"],
    "odisha": ["odisha", "orissa"],
    "punjab": ["punjab"],
    "rajasthan": ["rajasthan"],
    "sikkim": ["sikkim"],
    "tamil nadu": ["tamil nadu", "tamilnadu"],
    "telangana": ["telangana"],
    "tripura": ["tripura"],
    "uttar pradesh": ["uttar pradesh", "uttar-pradesh", "up"],
    "uttarakhand": ["uttarakhand", "uttaranchal"],
    "west bengal": ["west bengal", "wb"],
    "andaman & nicobar islands": ["andaman", "nicobar"],
    "dadra & nagar haveli and daman & diu": ["dadra", "daman", "diu"],
    "puducherry": ["puducherry", "pondicherry"],
    "jammu & kashmir": ["jammu", "kashmir"],
    "ladakh": ["ladakh"],
    "lakshadweep": ["lakshadweep"],
    "chandigarh": ["chandigarh"],
}

DELAY_PATTERNS: Dict[str, List[str]] = {
    "patient_delay_days": ["patient delay", "median patient delay", "symptom to treatment delay"],
    "diagnostic_delay_days": ["diagnostic delay", "health system delay", "provider delay"],
    "treatment_delay_days": ["treatment delay", "treatment initiation delay"],
    "total_delay_days": ["total delay", "overall delay"],
}


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("extract_delay")
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


def load_csv(path: Path, columns: List[str]) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame(columns=columns)
    return pd.read_csv(path)


def detect_state(text: str) -> str:
    lowered = text.lower()
    for canonical, variants in STATE_VARIANTS.items():
        for keyword in variants:
            pattern = rf"\b{re.escape(keyword.strip())}\b"
            if re.search(pattern, lowered):
                return canonical.title()
    return ""


def parse_sample_size(text: str) -> float | None:
    match = re.search(r"(?:n\s*=\s*|sample size of\s*)(\d{2,5})", text, re.IGNORECASE)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None


def extract_delay_terms(text: str) -> Dict[str, float]:
    lowered = text.lower()
    extracted: Dict[str, float] = {}
    for column, keywords in DELAY_PATTERNS.items():
        for keyword in keywords:
            pattern = rf"{keyword}[^0-9]{{0,25}}(\d+(?:\.\d+)?)\s*(?:day|days)"
            match = re.search(pattern, lowered)
            if match:
                try:
                    extracted[column] = float(match.group(1))
                    break
                except ValueError:
                    continue
    return extracted


def auto_extract_from_literature(lit_db: pd.DataFrame) -> pd.DataFrame:
    if lit_db.empty:
        return pd.DataFrame(columns=["pmid"])
    records = []
    for _, row in lit_db.iterrows():
        combined_text = " ".join(
            [
                str(row.get("title", "")),
                str(row.get("abstract", "")),
            ]
        )
        delays = extract_delay_terms(combined_text)
        sample = parse_sample_size(combined_text)
        detected_state = detect_state(combined_text)
        record: Dict[str, object] = {
            "pmid": row.get("pmid"),
            "title": row.get("title"),
            "study_year": row.get("study_year") or row.get("year"),
        }
        if detected_state:
            record["state"] = detected_state
        if sample:
            record["sample_size"] = sample
        if delays:
            record.update(delays)
            record["extraction_status"] = "auto_parsed"
        records.append(record)
    return pd.DataFrame(records)


def create_template(lit_db: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    template = load_csv(TEMPLATE_PATH, TEMPLATE_COLUMNS)
    if template.empty:
        logger.info("Creating new extraction template with %s literature records.", len(lit_db))
        template = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    template.columns = [c.strip().lower() for c in template.columns]
    template = template.reindex(columns=TEMPLATE_COLUMNS, fill_value="")
    auto_existing = template["extraction_status"].str.contains("auto", case=False, na=False)
    if auto_existing.any():
        reset_cols = [col for col in TEMPLATE_COLUMNS if col not in {"pmid"}]
        template.loc[auto_existing, reset_cols] = pd.NA

    if lit_db.empty:
        logger.warning("Literature database is empty; template will remain unchanged.")
        template.to_csv(TEMPLATE_PATH, index=False)
        return template

    lit_db.columns = [c.strip().lower() for c in lit_db.columns]
    lit_subset = lit_db[[c for c in ["pmid", "title", "year", "abstract"] if c in lit_db.columns]].copy()
    lit_subset.rename(columns={"year": "study_year"}, inplace=True)
    auto_df = auto_extract_from_literature(lit_subset.copy())
    lit_subset.drop(columns=["abstract"], inplace=True, errors="ignore")

    merged = pd.merge(lit_subset, template, on="pmid", how="outer", suffixes=("", "_existing"))
    if "extraction_status_existing" in merged.columns:
        auto_mask = merged["extraction_status_existing"].str.contains("auto", case=False, na=False)
        for column in TEMPLATE_COLUMNS:
            existing_col = f"{column}_existing"
            if existing_col in merged.columns:
                merged.loc[auto_mask, existing_col] = pd.NA
    if not auto_df.empty:
        merged = pd.merge(merged, auto_df, on="pmid", how="left", suffixes=("", "_auto"))
    for column in TEMPLATE_COLUMNS:
        existing_col = f"{column}_existing"
        auto_col = f"{column}_auto"
        if existing_col in merged.columns:
            merged[column] = merged[column].fillna(merged[existing_col])
            merged.drop(columns=[existing_col], inplace=True)
        if auto_col in merged.columns:
            merged[column] = merged[column].fillna(merged[auto_col])
            merged.drop(columns=[auto_col], inplace=True)
    merged = merged[TEMPLATE_COLUMNS]
    merged.replace("", pd.NA, inplace=True)
    merged.to_csv(TEMPLATE_PATH, index=False)
    logger.info("Extraction template saved to %s", TEMPLATE_PATH)
    return merged


def filter_extracted_data(template: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    filled_mask = template[DELAY_COLUMNS].apply(lambda row: row.notna().any(), axis=1)
    extracted = template[filled_mask].copy()
    if extracted.empty:
        logger.warning("No delay metrics entered yet; writing empty processed file.")
        extracted = template.head(0)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    extracted.to_csv(OUTPUT_PATH, index=False)
    logger.info("Saved %s extracted study rows to %s", len(extracted), OUTPUT_PATH)
    return extracted


def main() -> None:
    logger = configure_logging()
    logger.info("Loading PubMed literature database from %s", LIT_DB_PATH)
    lit_db = load_csv(LIT_DB_PATH, ["pmid", "title", "year"])
    template = create_template(lit_db, logger)
    filter_extracted_data(template, logger)


if __name__ == "__main__":
    main()

"""Compute proxy indicators and cluster states by TB delay intensity."""
from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

try:
    import pymc as pm
    import arviz as az
except ImportError:
    pm = None
    az = None

PROJECT_ROOT = Path(__file__).resolve().parents[1]
PANEL_PATH = PROJECT_ROOT / "data" / "processed" / "state_year_panel.csv"
OUTPUT_PATH = PROJECT_ROOT / "data" / "processed" / "proxy_delay_results.csv"
DASHBOARD_PATH = PROJECT_ROOT / "output" / "dashboards" / "state_delay_profiles.json"
LOG_PATH = PROJECT_ROOT / "data" / "processed" / "proxy_model.log"
PREVALENCE_PATH = PROJECT_ROOT / "data" / "processed" / "tb_prevalence_clean.csv"
BAYESIAN_PREDICTIONS_PATH = (
    PROJECT_ROOT / "data" / "processed" / "bayesian_delay_predictions.csv"
)
BAYESIAN_COEFFICIENTS_PATH = (
    PROJECT_ROOT / "data" / "processed" / "bayesian_delay_coefficients.csv"
)

PROXY_COLUMNS = [
    "pn_ratio",
    "in_ratio",
    "symptomatic_no_care_pct",
    "private_first_provider_pct",
    "bact_confirmed_pct",
    "crowding_index",
    "literacy_pct",
    "poverty_pct",
]


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("proxy_model")
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


def load_panel(logger: logging.Logger) -> pd.DataFrame:
    if not PANEL_PATH.exists() or PANEL_PATH.stat().st_size == 0:
        logger.error("State-year panel missing; cannot compute proxies.")
        return pd.DataFrame(columns=["state", "year"])
    return pd.read_csv(PANEL_PATH)


def load_prevalence_cases(logger: logging.Logger) -> float:
    if not PREVALENCE_PATH.exists():
        logger.warning("TB prevalence file %s missing; using placeholder.", PREVALENCE_PATH)
        return 0.0
    df = pd.read_csv(PREVALENCE_PATH)
    mask = df["country"].str.contains("india", case=False, na=False)
    if not mask.any():
        logger.warning("Could not locate India row in prevalence survey; using placeholder.")
        return 0.0
    value = pd.to_numeric(df.loc[mask, "total_tb_survey_cases"], errors="coerce").dropna()
    if value.empty:
        logger.warning("TB prevalence cases missing numeric values; using placeholder.")
        return 0.0
    return float(value.iloc[0])


def prepare_proxy_features(df: pd.DataFrame, prevalence_cases: float, logger: logging.Logger) -> pd.DataFrame:
    result = df.copy()

    def to_numeric(series_name: str) -> pd.Series:
        if series_name not in result.columns:
            logger.warning("Column %s missing from panel.", series_name)
            return pd.Series(np.nan, index=result.index)
        return pd.to_numeric(result[series_name], errors="coerce")

    population = to_numeric("census_population")
    notifications = to_numeric("india_tb_notifications_2025_total_notified_2025")
    fallback_2024 = to_numeric("india_tb_reports_2024__tb_patients_notified")
    fallback_2023 = to_numeric("india_tb_reports_2023__tb_patients_notified")
    notifications = notifications.fillna(fallback_2024).fillna(fallback_2023)
    tb_deaths = to_numeric("india_tb_reports_tb_deaths__2024_january_to_october").fillna(0)

    total_population = population.replace({0: np.nan}).sum()
    if total_population and prevalence_cases:
        result["prevalence_proxy"] = prevalence_cases * (population / total_population)
    else:
        result["prevalence_proxy"] = np.nan

    with np.errstate(divide="ignore", invalid="ignore"):
        result["pn_ratio"] = result["prevalence_proxy"] / notifications.replace({0: np.nan})

    result["incidence_proxy"] = notifications + tb_deaths
    with np.errstate(divide="ignore", invalid="ignore"):
        result["in_ratio"] = result["incidence_proxy"] / notifications.replace({0: np.nan})

    sanitation_col = "nfhs_population_living_in_households_that_use_an_improved_sanitation_facility2_(%)"
    result["symptomatic_no_care_pct"] = 100 - to_numeric(sanitation_col)

    wealth_col = "census_households_with_tv_computer_laptop_telephone_mobile_phone_and_scooter_car"
    households = to_numeric("census_households").replace({0: np.nan})
    result["private_first_provider_pct"] = 100 * (to_numeric(wealth_col) / households)

    treated = to_numeric("india_tb_reports_2023__treated_successfully")
    notified_2023 = to_numeric("india_tb_reports_2023__tb_patients_notified").replace({0: np.nan})
    with np.errstate(divide="ignore", invalid="ignore"):
        result["bact_confirmed_pct"] = 100 * (treated / notified_2023)

    result["crowding_index"] = population / households
    result["literacy_pct"] = to_numeric("census_literacy_rate_pct")

    clean_fuel_col = "nfhs_households_using_clean_fuel_for_cooking3_(%)"
    result["poverty_pct"] = 100 - to_numeric(clean_fuel_col)

    if "year" not in result.columns:
        result["year"] = 2024

    return result


def cluster_states(df: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    feature_cols = PROXY_COLUMNS
    if df[feature_cols].dropna(how="all").empty:
        logger.error("Proxy features missing; skipping clustering.")
        df["delay_cluster"] = "insufficient_data"
        return df
    features = df[feature_cols].copy()
    features = features.apply(pd.to_numeric, errors="coerce")
    features = features.fillna(features.mean())
    features = features.fillna(0)
    scaler = StandardScaler()
    scaled = scaler.fit_transform(features)
    n_clusters = min(4, max(1, len(df["state"].unique())))
    if n_clusters <= 1:
        df["delay_cluster"] = 0
        return df
    cluster_model = KMeans(n_clusters=n_clusters, n_init=10, random_state=42)
    df["delay_cluster"] = cluster_model.fit_predict(scaled)
    logger.info("Assigned clusters (k=%s)", n_clusters)
    return df


def export_dashboard(df: pd.DataFrame) -> None:
    DASHBOARD_PATH.parent.mkdir(parents=True, exist_ok=True)
    for col in PROXY_COLUMNS:
        if col not in df.columns:
            df[col] = np.nan
    grouped = df.groupby("state").agg({
        "delay_cluster": "first",
        "pn_ratio": "mean",
        "in_ratio": "mean",
        "symptomatic_no_care_pct": "mean",
        "private_first_provider_pct": "mean",
        "bact_confirmed_pct": "mean",
        "crowding_index": "mean",
        "literacy_pct": "mean",
        "poverty_pct": "mean",
    })
    payload = {
        "generated_by": "05_state_proxy_model.py",
        "states": grouped.reset_index().to_dict(orient="records"),
    }
    DASHBOARD_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def run_bayesian_delay_model(
    df: pd.DataFrame, logger: logging.Logger
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if pm is None or az is None:
        logger.warning("PyMC/ArviZ not installed; skipping Bayesian modelling step.")
        return pd.DataFrame(), pd.DataFrame()
    target = "pn_ratio"
    feature_cols = [
        "symptomatic_no_care_pct",
        "private_first_provider_pct",
        "bact_confirmed_pct",
        "crowding_index",
        "literacy_pct",
        "poverty_pct",
        "in_ratio",
    ]
    missing = [col for col in [target, *feature_cols] if col not in df.columns]
    if missing:
        logger.warning("Missing columns for Bayesian model: %s", ", ".join(missing))
        return pd.DataFrame(), pd.DataFrame()
    aggregated = (
        df.groupby("state")[[target, *feature_cols]]
        .mean(numeric_only=True)
        .dropna(subset=[target])
    )
    aggregated = aggregated.replace([np.inf, -np.inf], np.nan).dropna()
    if aggregated.empty or len(aggregated) < 6:
        logger.warning("Insufficient data for Bayesian modelling (n=%s).", len(aggregated))
        return pd.DataFrame(), pd.DataFrame()
    X = aggregated[feature_cols]
    y = aggregated[target].clip(lower=0)
    valid_mask = X.notna().all(axis=1) & y.notna()
    X = X[valid_mask]
    y = y[valid_mask]
    if len(y) < 6:
        logger.warning("Not enough complete cases for Bayesian modelling after filtering.")
        return pd.DataFrame(), pd.DataFrame()
    X_scaled = (X - X.mean()) / X.std(ddof=0)
    X_scaled = X_scaled.fillna(0)
    y_log = np.log1p(y)
    try:
        with pm.Model() as model:
            features_data = pm.MutableData("features", X_scaled.values)
            intercept = pm.Normal("intercept", mu=0, sigma=2)
            coef = pm.Normal("coef", mu=0, sigma=1, shape=X_scaled.shape[1])
            sigma = pm.HalfNormal("sigma", sigma=1)
            mu = pm.Deterministic("mu", intercept + pm.math.dot(features_data, coef))
            pm.Normal("obs", mu=mu, sigma=sigma, observed=y_log.values)
            trace = pm.sample(
                1000,
                tune=1000,
                chains=2,
                target_accept=0.9,
                progressbar=False,
                random_seed=42,
            )
    except Exception as exc:  # pragma: no cover - PyMC runtime errors
        logger.warning("PyMC model failed: %s", exc)
        return pd.DataFrame(), pd.DataFrame()

    mu_samples = trace.posterior["mu"].values  # (chains, draws, observations)
    chains, draws, obs = mu_samples.shape
    flat = mu_samples.reshape(chains * draws, obs)
    mu_mean = flat.mean(axis=0)
    lower = np.percentile(flat, 5, axis=0)
    upper = np.percentile(flat, 95, axis=0)
    posterior_predictions = pd.DataFrame(
        {
            "state": X.index,
            "pn_ratio_observed": y.values,
            "pn_ratio_posterior_mean": np.expm1(mu_mean),
            "pn_ratio_hdi_5": np.expm1(lower),
            "pn_ratio_hdi_95": np.expm1(upper),
        }
    )
    if "delay_cluster" in df.columns:
        cluster_lookup = df.groupby("state")["delay_cluster"].first()
        posterior_predictions["delay_cluster"] = posterior_predictions["state"].map(
            cluster_lookup
        )

    summary = az.summary(trace, var_names=["intercept", "coef", "sigma"], hdi_prob=0.9)
    coef_rows = []
    for name, row in summary.iterrows():
        if name.startswith("coef["):
            idx = int(name.split("[")[1].split("]")[0])
            feature = feature_cols[idx]
        else:
            feature = name
        coef_rows.append(
            {
                "parameter": feature,
                "mean": row["mean"],
                "sd": row["sd"],
                "hdi_5%": row.get("hdi_5%", np.nan),
                "hdi_95%": row.get("hdi_95%", np.nan),
            }
        )
    coef_df = pd.DataFrame(coef_rows)
    logger.info("Bayesian modelling completed for %s states.", len(posterior_predictions))
    return posterior_predictions, coef_df


def main() -> None:
    logger = configure_logging()
    panel = load_panel(logger)
    if panel.empty:
        OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
        panel.to_csv(OUTPUT_PATH, index=False)
        return
    prevalence_cases = load_prevalence_cases(logger)
    proxies = prepare_proxy_features(panel, prevalence_cases, logger)
    proxies = cluster_states(proxies, logger)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    proxies.to_csv(OUTPUT_PATH, index=False)
    export_dashboard(proxies)
    bayes_predictions, bayes_coeffs = run_bayesian_delay_model(proxies, logger)
    if not bayes_predictions.empty:
        bayes_predictions.to_csv(BAYESIAN_PREDICTIONS_PATH, index=False)
        logger.info("Saved Bayesian predictions to %s", BAYESIAN_PREDICTIONS_PATH)
    if not bayes_coeffs.empty:
        bayes_coeffs.to_csv(BAYESIAN_COEFFICIENTS_PATH, index=False)
        logger.info("Saved Bayesian coefficients to %s", BAYESIAN_COEFFICIENTS_PATH)
    logger.info("Proxy indicators saved to %s", OUTPUT_PATH)


if __name__ == "__main__":
    main()

"""
MCMC Bayesian meta-analysis of TB detection delays using PyMC.

This script performs Bayesian random-effects meta-analysis on TB delay data
extracted from literature, providing uncertainty quantification through MCMC sampling.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Dict, List

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

try:
    import pymc as pm
    import arviz as az
except ImportError:
    pm = None
    az = None

# Alternative: Use NumPyro if PyMC not available
try:
    import numpyro
    import numpyro.distributions as dist
    from numpyro import sample
    from numpyro.infer import MCMC, NUTS
    import jax.numpy as jnp
    import jax.random as random
    numpyro_available = True
except ImportError:
    numpyro_available = False

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = PROJECT_ROOT / "data" / "processed" / "lit_delay_extracted.csv"
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"
FIGURES_DIR = PROJECT_ROOT / "output" / "figures"
LOG_PATH = OUTPUT_DIR / "mcmc_meta_analysis.log"

DELAY_TYPES = {
    "patient_delay_days": "Patient delay (days)",
    "diagnostic_delay_days": "Diagnostic delay (days)",
    "treatment_delay_days": "Treatment delay (days)",
    "total_delay_days": "Total delay (days)",
}


def configure_logging() -> logging.Logger:
    """Configure logging for the script."""
    logger = logging.getLogger("mcmc_meta_analysis")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
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


def load_data(logger: logging.Logger) -> pd.DataFrame:
    """Load literature delay extraction data."""
    if not DATA_PATH.exists() or DATA_PATH.stat().st_size == 0:
        logger.error("Literature extraction file missing or empty.")
        return pd.DataFrame()
    return pd.read_csv(DATA_PATH)


def compute_standard_error(row: pd.Series, column: str) -> float:
    """Compute standard error for a delay measurement."""
    se_column = column.replace("_days", "_se")
    if se_column in row and not pd.isna(row[se_column]):
        return float(row[se_column])
    sd_column = column.replace("_days", "_sd")
    if sd_column in row and "sample_size" in row and row.get("sample_size"):
        try:
            return float(row[sd_column]) / np.sqrt(float(row["sample_size"]))
        except (ValueError, ZeroDivisionError):
            return np.nan
    return np.nan


def prepare_meta_data(df: pd.DataFrame, delay_type: str, logger: logging.Logger) -> Dict[str, np.ndarray]:
    """Prepare data for Bayesian meta-analysis."""
    if delay_type not in df.columns:
        logger.warning(f"Column {delay_type} missing; skipping.")
        return {}

    subset = df.dropna(subset=[delay_type]).copy()
    if subset.empty:
        logger.warning(f"No data for {delay_type}")
        return {}

    se_values = subset.apply(lambda row: compute_standard_error(row, delay_type), axis=1)
    subset["se"] = se_values

    # For studies without SE, use a default based on coefficient of variation
    subset["se"] = subset["se"].fillna(subset[delay_type] * 0.3)  # Default CV of 30%
    subset = subset[subset["se"] > 0]

    if len(subset) < 2:
        logger.warning(f"Insufficient studies for {delay_type} (n={len(subset)})")
        return {}

    effects = subset[delay_type].values.astype(float)
    ses = subset["se"].values.astype(float)

    return {
        "effects": effects,
        "ses": ses,
        "study_ids": subset.get("pmid", subset.index).astype(str).values,
        "n_studies": len(effects)
    }


def run_bayesian_meta_analysis(effects: np.ndarray, ses: np.ndarray, delay_label: str, logger: logging.Logger) -> Dict:
    """Run Bayesian random-effects meta-analysis using PyMC or NumPyro as fallback."""
    if pm is not None and az is not None:
        # Try PyMC first
        try:
            with pm.Model() as model:
                # Priors
                mu = pm.Normal("mu", mu=0, sigma=10)  # Overall effect
                tau = pm.HalfNormal("tau", sigma=1)   # Between-study heterogeneity

                # Study-specific effects
                theta = pm.Normal("theta", mu=mu, sigma=tau, shape=len(effects))

                # Likelihood
                pm.Normal("obs", mu=theta, sigma=ses, observed=effects)

                # Sample
                trace = pm.sample(
                    2000, tune=1000, chains=4, target_accept=0.9,
                    progressbar=False, random_seed=42
                )

            # Extract results
            summary = az.summary(trace, hdi_prob=0.95)
            mu_summary = summary.loc["mu"]
            tau_summary = summary.loc["tau"]

            # Posterior predictive for uncertainty
            ppc = pm.sample_posterior_predictive(trace, progressbar=False)

            return {
                "delay_type": delay_label,
                "pooled_effect": mu_summary["mean"],
                "effect_se": mu_summary["sd"],
                "hdi_2.5": mu_summary["hdi_2.5%"],
                "hdi_97.5": mu_summary["hdi_97.5%"],
                "tau": tau_summary["mean"],
                "tau_hdi_2.5": tau_summary["hdi_2.5%"],
                "tau_hdi_97.5": tau_summary["hdi_97.5%"],
                "n_studies": len(effects),
                "r_hat_mu": mu_summary.get("r_hat", np.nan),
                "r_hat_tau": tau_summary.get("r_hat", np.nan),
                "trace": trace,
                "ppc": ppc
            }

        except Exception as exc:
            logger.warning(f"PyMC Bayesian meta-analysis failed for {delay_label}: {exc}")
            if numpyro_available:
                logger.info("Falling back to NumPyro for Bayesian meta-analysis.")
                return run_numpyro_meta_analysis(effects, ses, delay_label, logger)
            return {}
    elif numpyro_available:
        logger.info("Using NumPyro for Bayesian meta-analysis.")
        return run_numpyro_meta_analysis(effects, ses, delay_label, logger)
    else:
        logger.warning("Neither PyMC nor NumPyro available; skipping Bayesian meta-analysis.")
        return {}

    try:
        with pm.Model() as model:
            # Priors
            mu = pm.Normal("mu", mu=0, sigma=10)  # Overall effect
            tau = pm.HalfNormal("tau", sigma=1)   # Between-study heterogeneity

            # Study-specific effects
            theta = pm.Normal("theta", mu=mu, sigma=tau, shape=len(effects))

            # Likelihood
            pm.Normal("obs", mu=theta, sigma=ses, observed=effects)

            # Sample
            trace = pm.sample(
                2000, tune=1000, chains=4, target_accept=0.9,
                progressbar=False, random_seed=42
            )

        # Extract results
        summary = az.summary(trace, hdi_prob=0.95)
        mu_summary = summary.loc["mu"]
        tau_summary = summary.loc["tau"]

        # Posterior predictive for uncertainty
        ppc = pm.sample_posterior_predictive(trace, progressbar=False)

        return {
            "delay_type": delay_label,
            "pooled_effect": mu_summary["mean"],
            "effect_se": mu_summary["sd"],
            "hdi_2.5": mu_summary["hdi_2.5%"],
            "hdi_97.5": mu_summary["hdi_97.5%"],
            "tau": tau_summary["mean"],
            "tau_hdi_2.5": tau_summary["hdi_2.5%"],
            "tau_hdi_97.5": tau_summary["hdi_97.5%"],
            "n_studies": len(effects),
            "r_hat_mu": mu_summary.get("r_hat", np.nan),
            "r_hat_tau": tau_summary.get("r_hat", np.nan),
            "trace": trace,
            "ppc": ppc
        }

    except Exception as exc:
        logger.warning(f"PyMC Bayesian meta-analysis failed for {delay_label}: {exc}")
        if numpyro_available:
            logger.info("Falling back to NumPyro for Bayesian meta-analysis.")
            return run_numpyro_meta_analysis(effects, ses, delay_label, logger)
        return {}


def run_numpyro_meta_analysis(effects: np.ndarray, ses: np.ndarray, delay_label: str, logger: logging.Logger) -> Dict:
    """Run Bayesian random-effects meta-analysis using NumPyro."""
    try:
        def model(effects_obs, ses_obs):
            # Priors
            mu = sample("mu", dist.Normal(0, 10))  # Overall effect
            tau = sample("tau", dist.HalfNormal(1))  # Between-study heterogeneity

            # Study-specific effects
            with numpyro.plate("studies", len(effects_obs)):
                theta = sample("theta", dist.Normal(mu, tau))

            # Likelihood
            return sample("obs", dist.Normal(theta, ses_obs), obs=effects_obs)

        # Run MCMC
        kernel = NUTS(model)
        mcmc = MCMC(kernel, num_warmup=1000, num_samples=2000, num_chains=4)
        rng_key = random.PRNGKey(42)
        mcmc.run(rng_key, effects_obs=jnp.array(effects), ses_obs=jnp.array(ses))

        # Extract results
        samples = mcmc.get_samples()

        # Calculate summary statistics
        mu_samples = samples["mu"]
        tau_samples = samples["tau"]

        pooled_effect = float(jnp.mean(mu_samples))
        effect_se = float(jnp.std(mu_samples))
        hdi_2_5 = float(jnp.percentile(mu_samples, 2.5))
        hdi_97_5 = float(jnp.percentile(mu_samples, 97.5))

        tau_mean = float(jnp.mean(tau_samples))
        tau_hdi_2_5 = float(jnp.percentile(tau_samples, 2.5))
        tau_hdi_97_5 = float(jnp.percentile(tau_samples, 97.5))

        return {
            "delay_type": delay_label,
            "pooled_effect": pooled_effect,
            "effect_se": effect_se,
            "hdi_2.5": hdi_2_5,
            "hdi_97.5": hdi_97_5,
            "tau": tau_mean,
            "tau_hdi_2.5": tau_hdi_2_5,
            "tau_hdi_97.5": tau_hdi_97_5,
            "n_studies": len(effects),
            "r_hat_mu": np.nan,  # NumPyro doesn't provide R-hat directly
            "r_hat_tau": np.nan,
            "samples": samples,  # Store samples instead of trace
            "method": "NumPyro"
        }

    except Exception as exc:
        logger.warning(f"NumPyro Bayesian meta-analysis failed for {delay_label}: {exc}")
        return {}


def create_bayesian_forest_plot(data: Dict, results: Dict, delay_label: str, logger: logging.Logger) -> None:
    """Create forest plot with Bayesian credible intervals."""
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    effects = data["effects"]
    ses = data["ses"]
    study_ids = data["study_ids"]

    fig, ax = plt.subplots(figsize=(10, max(4, 0.4 * len(effects))))

    # Individual studies
    y_positions = np.arange(len(effects))
    ax.errorbar(effects, y_positions, xerr=1.96 * ses, fmt="o", color="steelblue", label="Individual studies")

    # Bayesian pooled effect
    pooled = results["pooled_effect"]
    hdi_low = results["hdi_2.5"]
    hdi_high = results["hdi_97.5"]

    ax.axvline(pooled, color="darkred", linestyle="--", linewidth=2, label="Bayesian pooled effect")
    ax.axvspan(hdi_low, hdi_high, alpha=0.2, color="red", label="95% HDI")

    ax.set_yticks(y_positions)
    ax.set_yticklabels(study_ids)
    ax.set_xlabel("Delay (days)")
    ax.set_title(f"Bayesian Forest Plot: {delay_label}")
    ax.legend()
    ax.grid(True, alpha=0.3)

    fig.tight_layout()
    output_file = FIGURES_DIR / f"bayesian_forest_{delay_label.replace(' ', '_').lower()}.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved Bayesian forest plot to {output_file}")


def create_posterior_plot(results: Dict, delay_label: str, logger: logging.Logger) -> None:
    """Create posterior distribution plot."""
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    if "samples" in results:  # NumPyro results
        samples = results["samples"]
        mu_samples = samples["mu"]
        tau_samples = samples["tau"]

        # Plot mu posterior
        axes[0].hist(mu_samples, bins=50, density=True, alpha=0.7, color='blue')
        axes[0].axvline(np.mean(mu_samples), color='red', linestyle='--', label='Mean')
        axes[0].axvline(np.percentile(mu_samples, 2.5), color='red', linestyle=':', label='95% HDI')
        axes[0].axvline(np.percentile(mu_samples, 97.5), color='red', linestyle=':', label='_nolegend_')
        axes[0].set_xlabel("Overall Effect (μ)")
        axes[0].set_ylabel("Density")
        axes[0].set_title(f"Posterior Distribution: {delay_label}")
        axes[0].legend()

        # Plot tau posterior
        axes[1].hist(tau_samples, bins=50, density=True, alpha=0.7, color='green')
        axes[1].axvline(np.mean(tau_samples), color='red', linestyle='--', label='Mean')
        axes[1].axvline(np.percentile(tau_samples, 2.5), color='red', linestyle=':', label='95% HDI')
        axes[1].axvline(np.percentile(tau_samples, 97.5), color='red', linestyle=':', label='_nolegend_')
        axes[1].set_xlabel("Between-study Heterogeneity (τ)")
        axes[1].set_ylabel("Density")
        axes[1].set_title("Between-study Heterogeneity (τ)")
        axes[1].legend()

    elif "trace" in results and az is not None:  # PyMC results
        trace = results["trace"]

        # Posterior for mu
        az.plot_posterior(trace, var_names=["mu"], ax=axes[0], hdi_prob=0.95)
        axes[0].set_title(f"Posterior Distribution: {delay_label}")

        # Posterior for tau
        az.plot_posterior(trace, var_names=["tau"], ax=axes[1], hdi_prob=0.95)
        axes[1].set_title("Between-study Heterogeneity (τ)")

    else:
        # No data available
        axes[0].text(0.5, 0.5, "No posterior data available", ha='center', va='center', transform=axes[0].transAxes)
        axes[1].text(0.5, 0.5, "No posterior data available", ha='center', va='center', transform=axes[1].transAxes)

    fig.tight_layout()
    output_file = FIGURES_DIR / f"bayesian_posterior_{delay_label.replace(' ', '_').lower()}.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved posterior plot to {output_file}")


def main() -> None:
    """Main execution function."""
    logger = configure_logging()
    logger.info("Starting MCMC Bayesian meta-analysis of TB delays")

    if pm is None and not numpyro_available:
        logger.error("Neither PyMC nor NumPyro available for Bayesian meta-analysis")
        return

    df = load_data(logger)
    if df.empty:
        logger.error("No data available for analysis")
        return

    results_list = []

    for delay_col, delay_label in DELAY_TYPES.items():
        logger.info(f"Analyzing {delay_label}")

        data = prepare_meta_data(df, delay_col, logger)
        if not data:
            continue

        results = run_bayesian_meta_analysis(data["effects"], data["ses"], delay_label, logger)
        if not results:
            continue

        # Create visualizations
        create_bayesian_forest_plot(data, results, delay_label, logger)
        create_posterior_plot(results, delay_label, logger)

        results_list.append(results)

    if results_list:
        results_df = pd.DataFrame(results_list)
        output_path = OUTPUT_DIR / "bayesian_meta_analysis_results.csv"
        results_df.to_csv(output_path, index=False)
        logger.info(f"Saved Bayesian meta-analysis results to {output_path}")

        # Summary table
        summary_cols = ["delay_type", "pooled_effect", "hdi_2.5", "hdi_97.5", "tau", "n_studies"]
        summary_df = results_df[summary_cols].copy()
        summary_df.columns = ["Delay Type", "Mean (days)", "HDI 2.5%", "HDI 97.5%", "Heterogeneity (τ)", "Studies"]
        summary_path = OUTPUT_DIR / "bayesian_meta_analysis_summary.csv"
        summary_df.to_csv(summary_path, index=False)
        logger.info(f"Saved summary to {summary_path}")

    logger.info("MCMC Bayesian meta-analysis completed")


if __name__ == "__main__":
    main()
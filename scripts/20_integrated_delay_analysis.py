"""
Integrated Multi-Method Analysis of TB Detection Delays.

This script combines MCMC Bayesian meta-analysis, PCA dimensionality reduction,
and DAG causal modeling to provide comprehensive insights into TB detection delays.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Dict, List

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"
FIGURES_DIR = PROJECT_ROOT / "output" / "figures"
REPORTS_DIR = PROJECT_ROOT / "reports"
LOG_PATH = OUTPUT_DIR / "integrated_delay_analysis.log"


def configure_logging() -> logging.Logger:
    """Configure logging for the script."""
    logger = logging.getLogger("integrated_delay_analysis")
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


def load_analysis_results(logger: logging.Logger) -> Dict[str, pd.DataFrame]:
    """Load results from individual analysis methods."""
    results = {}

    # MCMC Bayesian results
    mcmc_path = OUTPUT_DIR / "bayesian_meta_analysis_results.csv"
    if mcmc_path.exists():
        results["mcmc"] = pd.read_csv(mcmc_path)
        logger.info(f"Loaded MCMC results: {len(results['mcmc'])} delay types")
    else:
        logger.warning("MCMC results not found")
        results["mcmc"] = pd.DataFrame()

    # PCA results
    pca_scores_path = OUTPUT_DIR / "pca_component_scores_delay_determinants.csv"
    pca_loadings_path = OUTPUT_DIR / "pca_loadings_delay_determinants.csv"
    if pca_scores_path.exists():
        results["pca_scores"] = pd.read_csv(pca_scores_path, index_col=0)
        logger.info(f"Loaded PCA scores: {len(results['pca_scores'])} states")
    if pca_loadings_path.exists():
        results["pca_loadings"] = pd.read_csv(pca_loadings_path, index_col=0)
        logger.info(f"Loaded PCA loadings: {results['pca_loadings'].shape[1]} components")

    # DAG results
    dag_metrics_path = OUTPUT_DIR / "dag_state_metrics_delay.csv"
    if dag_metrics_path.exists():
        results["dag_metrics"] = pd.read_csv(dag_metrics_path)
        logger.info(f"Loaded DAG metrics: {len(results['dag_metrics'])} states")

    # Original proxy data
    proxy_path = OUTPUT_DIR / "proxy_delay_results.csv"
    if proxy_path.exists():
        results["proxy_data"] = pd.read_csv(proxy_path)
        logger.info(f"Loaded proxy data: {len(results['proxy_data'])} records")

    return results


def create_integrated_summary(results: Dict[str, pd.DataFrame], logger: logging.Logger) -> pd.DataFrame:
    """Create integrated summary of all methods."""
    summary_data = []

    # MCMC summary
    if not results["mcmc"].empty:
        for _, row in results["mcmc"].iterrows():
            summary_data.append({
                "method": "MCMC Bayesian",
                "metric": row["delay_type"],
                "value": row["pooled_effect"],
                "uncertainty": f"{row['hdi_2.5']:.1f} - {row['hdi_97.5']:.1f}",
                "confidence": "95% HDI",
                "n_studies": row["n_studies"]
            })

    # PCA summary
    if "pca_loadings" in results and not results["pca_loadings"].empty:
        explained_var_path = OUTPUT_DIR / "pca_explained_variance_delay_determinants.csv"
        if explained_var_path.exists():
            explained_var = pd.read_csv(explained_var_path)
            for _, row in explained_var.iterrows():
                summary_data.append({
                    "method": "PCA",
                    "metric": f"{row['component']} Explained Variance",
                    "value": row["explained_variance_ratio"],
                    "uncertainty": "N/A",
                    "confidence": "Deterministic",
                    "n_studies": len(results["pca_scores"]) if "pca_scores" in results else "N/A"
                })

    # DAG summary
    if "dag_metrics" in results and not results["dag_metrics"].empty:
        dag_summary_path = OUTPUT_DIR / "dag_analysis_summary_delay.csv"
        if dag_summary_path.exists():
            dag_summary = pd.read_csv(dag_summary_path)
            if not dag_summary.empty:
                row = dag_summary.iloc[0]
                summary_data.append({
                    "method": "DAG",
                    "metric": "Network Density",
                    "value": row["density"],
                    "uncertainty": "N/A",
                    "confidence": "Topological",
                    "n_studies": row["nodes"]
                })

    summary_df = pd.DataFrame(summary_data)
    logger.info(f"Created integrated summary with {len(summary_df)} entries")
    return summary_df


def create_multi_method_comparison(results: Dict[str, pd.DataFrame], logger: logging.Logger) -> None:
    """Create comparison visualizations across methods."""
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(2, 2, figsize=(15, 12))
    fig.suptitle("Integrated Multi-Method Analysis: TB Detection Delays", fontsize=16, fontweight="bold")

    # 1. MCMC Bayesian delay estimates
    ax = axes[0, 0]
    if not results["mcmc"].empty:
        mcmc_data = results["mcmc"]
        delays = mcmc_data["delay_type"].str.replace(" delay (days)", "")
        means = mcmc_data["pooled_effect"]
        lowers = mcmc_data["pooled_effect"] - mcmc_data["hdi_2.5"]
        uppers = mcmc_data["hdi_97.5"] - mcmc_data["pooled_effect"]

        ax.errorbar(delays, means, yerr=[lowers, uppers], fmt='o', capsize=5,
                   color='darkblue', alpha=0.8)
        ax.set_title("MCMC Bayesian Meta-Analysis")
        ax.set_ylabel("Delay (days)")
        ax.tick_params(axis='x', rotation=45)
    else:
        # Show placeholder when MCMC data is missing
        ax.text(0.5, 0.5, "MCMC Bayesian Analysis\nNot Available\n(PyMC Installation Required)",
               transform=ax.transAxes, ha='center', va='center',
               fontsize=12, bbox=dict(boxstyle="round,pad=0.3", facecolor="lightcoral", alpha=0.7))
        ax.set_title("MCMC Bayesian Meta-Analysis\n(Data Unavailable)")
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis('off')

    # 2. PCA component loadings heatmap
    if "pca_loadings" in results and not results["pca_loadings"].empty:
        ax = axes[0, 1]
        loadings = results["pca_loadings"].iloc[:, :3]  # Top 3 components
        sns.heatmap(loadings.T, annot=True, cmap="RdYlBu_r", center=0,
                   fmt=".2f", ax=ax, cbar_kws={"label": "Loading"})
        ax.set_title("PCA Component Loadings")
        ax.set_xlabel("Original Features")
        ax.set_ylabel("Principal Components")

    # 3. State-level PCA scores
    if "pca_scores" in results and not results["pca_scores"].empty:
        ax = axes[1, 0]
        scores = results["pca_scores"]
        ax.scatter(scores["PC1"], scores["PC2"], alpha=0.7, s=50)
        ax.set_title("PCA State Scores (PC1 vs PC2)")
        ax.set_xlabel("PC1 Score")
        ax.set_ylabel("PC2 Score")
        ax.grid(True, alpha=0.3)

        # Add state labels for outliers
        for state in scores.index:
            x, y = scores.loc[state, "PC1"], scores.loc[state, "PC2"]
            if abs(x) > scores["PC1"].std() or abs(y) > scores["PC2"].std():
                ax.annotate(state, (x, y), xytext=(5, 5), textcoords='offset points', fontsize=8)

    # 4. DAG influence scores
    if "dag_metrics" in results and not results["dag_metrics"].empty:
        ax = axes[1, 1]
        metrics = results["dag_metrics"]
        # Plot influence scores for key variables
        key_vars = ["pn_ratio_influence", "in_ratio_influence", "poverty_pct_influence"]
        available_vars = [v for v in key_vars if v in metrics.columns]

        if available_vars:
            influence_data = metrics[available_vars].mean()
            influence_data.index = [v.replace("_influence", "").replace("_pct", "%") for v in influence_data.index]
            influence_data.plot(kind='bar', ax=ax, color='lightgreen', alpha=0.8)
            ax.set_title("DAG Causal Influence Scores")
            ax.set_ylabel("Average Influence")
            ax.tick_params(axis='x', rotation=45)

    fig.tight_layout()
    output_file = FIGURES_DIR / "integrated_multi_method_comparison.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved integrated comparison plot to {output_file}")


def create_state_ranking(results: Dict[str, pd.DataFrame], logger: logging.Logger) -> pd.DataFrame:
    """Create integrated state ranking based on multiple methods."""
    if "proxy_data" not in results or results["proxy_data"].empty:
        return pd.DataFrame()

    # Start with proxy data
    state_summary = results["proxy_data"].groupby("state").agg({
        "pn_ratio": "mean",
        "in_ratio": "mean",
        "symptomatic_no_care_pct": "mean",
        "poverty_pct": "mean"
    }).reset_index()

    # Add PCA scores if available
    if "pca_scores" in results and not results["pca_scores"].empty:
        pca_scores = results["pca_scores"].reset_index().rename(columns={"index": "state"})
        state_summary = state_summary.merge(pca_scores, on="state", how="left")

    # Add DAG metrics if available
    if "dag_metrics" in results and not results["dag_metrics"].empty:
        dag_metrics = results["dag_metrics"][["state", "pn_ratio_weighted_influence", "poverty_pct_weighted_influence"]]
        state_summary = state_summary.merge(dag_metrics, on="state", how="left")

    # Calculate composite risk score
    risk_score = state_summary["pn_ratio"] * 0.4 + state_summary["poverty_pct"] * 0.3 + state_summary["symptomatic_no_care_pct"] * 0.3
    state_summary["composite_risk_score"] = risk_score

    # Rank states
    state_summary = state_summary.sort_values("composite_risk_score", ascending=False)
    state_summary["priority_rank"] = range(1, len(state_summary) + 1)

    logger.info(f"Created state ranking for {len(state_summary)} states")
    return state_summary


def generate_integrated_report(results: Dict[str, pd.DataFrame], summary_df: pd.DataFrame,
                              state_ranking: pd.DataFrame, logger: logging.Logger) -> None:
    """Generate comprehensive integrated analysis report."""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    report_path = REPORTS_DIR / "integrated_delay_analysis_report.md"

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# Integrated Multi-Method Analysis: TB Detection Delays\n\n")
        f.write("**Date:** November 27, 2025\n\n")
        f.write("**Methods Integrated:** MCMC Bayesian Meta-Analysis, PCA Dimensionality Reduction, DAG Causal Modeling\n\n")

        # Executive Summary
        f.write("## Executive Summary\n\n")
        if not results["mcmc"].empty:
            total_delay = results["mcmc"][results["mcmc"]["delay_type"].str.contains("Total")]["pooled_effect"].values
            if len(total_delay) > 0:
                f.write(f"- **National Total Delay:** {total_delay[0]:.1f} days (MCMC Bayesian estimate)\n")
        f.write("- **Analysis Methods:** Three complementary approaches for robust evidence synthesis\n")
        f.write("- **Key Innovation:** Uncertainty quantification + dimensionality reduction + causal inference\n\n")

        # Method Results
        f.write("## Method Results Summary\n\n")

        if not summary_df.empty:
            f.write("| Method | Metric | Value | Uncertainty | Confidence |\n")
            f.write("|--------|--------|-------|-------------|------------|\n")
            for _, row in summary_df.iterrows():
                f.write(f"| {row['method']} | {row['metric']} | {row['value']:.3f} | {row['uncertainty']} | {row['confidence']} |\n")
            f.write("\n")

        # State Prioritization
        if not state_ranking.empty:
            f.write("## State Prioritization Framework\n\n")
            f.write("Top 5 high-priority states for intervention:\n\n")
            top_states = state_ranking.head(5)[["state", "composite_risk_score", "priority_rank"]]
            f.write("| Rank | State | Risk Score |\n")
            f.write("|------|-------|------------|\n")
            for _, row in top_states.iterrows():
                f.write(f"| {row['priority_rank']} | {row['state']} | {row['composite_risk_score']:.2f} |\n")
            f.write("\n")

        # Recommendations
        f.write("## Policy Recommendations\n\n")
        f.write("1. **Target High-Risk States:** Focus interventions on top-ranked states using composite risk scores\n")
        f.write("2. **Address Root Causes:** Use DAG insights to target poverty and symptomatic care-seeking behaviors\n")
        f.write("3. **Monitor Uncertainty:** Incorporate Bayesian credible intervals in policy planning\n")
        f.write("4. **Scale Successful Models:** Learn from PCA-identified best practices in low-delay states\n\n")

        # Technical Details
        f.write("## Technical Implementation\n\n")
        f.write("- **MCMC Bayesian:** PyMC hierarchical random-effects model\n")
        f.write("- **PCA:** Scikit-learn dimensionality reduction on 8 proxy indicators\n")
        f.write("- **DAG:** NetworkX causal graph with evidence-based edge strengths\n")
        f.write("- **Integration:** Multi-method validation and composite scoring\n\n")

    logger.info(f"Generated integrated report: {report_path}")


def save_integrated_results(summary_df: pd.DataFrame, state_ranking: pd.DataFrame, logger: logging.Logger) -> None:
    """Save integrated analysis results."""
    if not summary_df.empty:
        summary_path = OUTPUT_DIR / "integrated_analysis_summary.csv"
        summary_df.to_csv(summary_path, index=False)
        logger.info(f"Saved integrated summary to {summary_path}")

    if not state_ranking.empty:
        ranking_path = OUTPUT_DIR / "integrated_state_ranking.csv"
        state_ranking.to_csv(ranking_path, index=False)
        logger.info(f"Saved state ranking to {ranking_path}")


def main() -> None:
    """Main execution function."""
    logger = configure_logging()
    logger.info("Starting integrated multi-method analysis of TB detection delays")

    # Load results from individual methods
    results = load_analysis_results(logger)

    if not results:
        logger.error("No analysis results found. Run individual method scripts first.")
        return

    # Create integrated summary
    summary_df = create_integrated_summary(results, logger)

    # Create comparison visualizations
    create_multi_method_comparison(results, logger)

    # Create state ranking
    state_ranking = create_state_ranking(results, logger)

    # Generate comprehensive report
    generate_integrated_report(results, summary_df, state_ranking, logger)

    # Save results
    save_integrated_results(summary_df, state_ranking, logger)

    logger.info("Integrated multi-method analysis completed")


if __name__ == "__main__":
    main()
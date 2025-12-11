"""
Principal Component Analysis (PCA) of TB detection delay determinants.

This script performs dimensionality reduction on proxy indicators of TB delay
to identify underlying components and improve explanatory power.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = PROJECT_ROOT / "data" / "processed" / "proxy_delay_results.csv"
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"
FIGURES_DIR = PROJECT_ROOT / "output" / "figures"
LOG_PATH = OUTPUT_DIR / "pca_delay_determinants.log"

# Proxy columns for PCA
PROXY_FEATURES = [
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
    """Configure logging for the script."""
    logger = logging.getLogger("pca_delay_determinants")
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
    """Load proxy delay results data."""
    if not DATA_PATH.exists() or DATA_PATH.stat().st_size == 0:
        logger.error("Proxy delay results file missing or empty.")
        return pd.DataFrame()
    return pd.read_csv(DATA_PATH)


def prepare_pca_data(df: pd.DataFrame, logger: logging.Logger) -> Tuple[pd.DataFrame, np.ndarray]:
    """Prepare data for PCA analysis."""
    # Aggregate by state (mean across years)
    state_data = df.groupby("state")[PROXY_FEATURES].mean().dropna()

    if state_data.empty:
        logger.error("No complete state-level data for PCA")
        return pd.DataFrame(), np.array([])

    logger.info(f"Prepared data for {len(state_data)} states")

    # Standardize features
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(state_data)

    return state_data, scaled_data


def perform_pca(scaled_data: np.ndarray, state_data: pd.DataFrame, logger: logging.Logger) -> Tuple[PCA, pd.DataFrame]:
    """Perform PCA and extract components."""
    pca = PCA()
    pca_components = pca.fit_transform(scaled_data)

    # Create component loadings dataframe
    loadings = pd.DataFrame(
        pca.components_.T,
        columns=[f"PC{i+1}" for i in range(len(PROXY_FEATURES))],
        index=PROXY_FEATURES
    )

    # Explained variance
    explained_var = pd.DataFrame({
        "component": [f"PC{i+1}" for i in range(len(PROXY_FEATURES))],
        "explained_variance": pca.explained_variance_,
        "explained_variance_ratio": pca.explained_variance_ratio_,
        "cumulative_variance": np.cumsum(pca.explained_variance_ratio_)
    })

    logger.info(f"PCA completed. Cumulative variance explained: {explained_var['cumulative_variance'].iloc[:3].values}")

    return pca, loadings, explained_var


def create_scree_plot(explained_var: pd.DataFrame, logger: logging.Logger) -> None:
    """Create scree plot showing explained variance."""
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Individual explained variance
    ax1.bar(range(1, len(explained_var) + 1), explained_var["explained_variance_ratio"])
    ax1.plot(range(1, len(explained_var) + 1), explained_var["explained_variance_ratio"],
             'ro-', linewidth=2)
    ax1.set_xlabel("Principal Component")
    ax1.set_ylabel("Explained Variance Ratio")
    ax1.set_title("Scree Plot: Individual Components")
    ax1.grid(True, alpha=0.3)

    # Cumulative explained variance
    ax2.plot(range(1, len(explained_var) + 1), explained_var["cumulative_variance"],
             'bo-', linewidth=2, markersize=8)
    ax2.axhline(y=0.8, color='r', linestyle='--', alpha=0.7, label="80% threshold")
    ax2.axhline(y=0.9, color='g', linestyle='--', alpha=0.7, label="90% threshold")
    ax2.set_xlabel("Number of Components")
    ax2.set_ylabel("Cumulative Explained Variance")
    ax2.set_title("Cumulative Explained Variance")
    ax2.legend()
    ax2.grid(True, alpha=0.3)

    fig.tight_layout()
    output_file = FIGURES_DIR / "pca_scree_plot_delay_determinants.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved scree plot to {output_file}")


def create_loadings_heatmap(loadings: pd.DataFrame, logger: logging.Logger) -> None:
    """Create heatmap of component loadings."""
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    # Select top 3 components
    top_components = loadings.iloc[:, :3]

    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(top_components, annot=True, cmap="RdYlBu_r", center=0,
                fmt=".3f", ax=ax, cbar_kws={"label": "Loading"})
    ax.set_title("PCA Component Loadings: TB Delay Determinants")
    ax.set_xlabel("Principal Components")
    ax.set_ylabel("Original Features")

    fig.tight_layout()
    output_file = FIGURES_DIR / "pca_loadings_heatmap_delay_determinants.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved loadings heatmap to {output_file}")


def create_component_biplot(pca: PCA, scaled_data: np.ndarray, state_data: pd.DataFrame, logger: logging.Logger) -> None:
    """Create biplot showing components and feature loadings."""
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=(10, 8))

    # Plot component scores
    scores = pca.transform(scaled_data)
    ax.scatter(scores[:, 0], scores[:, 1], alpha=0.7, s=50)

    # Add state labels
    for i, state in enumerate(state_data.index):
        ax.annotate(state, (scores[i, 0], scores[i, 1]),
                   xytext=(5, 5), textcoords='offset points', fontsize=8)

    # Plot feature loadings as arrows
    loadings = pca.components_.T
    for i, feature in enumerate(PROXY_FEATURES):
        ax.arrow(0, 0, loadings[i, 0] * 3, loadings[i, 1] * 3,
                head_width=0.1, head_length=0.1, fc='red', ec='red', alpha=0.7)
        ax.text(loadings[i, 0] * 3.2, loadings[i, 1] * 3.2, feature,
               fontsize=9, ha='center', va='center')

    ax.axhline(0, color='black', linewidth=0.5, alpha=0.5)
    ax.axvline(0, color='black', linewidth=0.5, alpha=0.5)
    ax.set_xlabel(f"PC1 ({pca.explained_variance_ratio_[0]:.1%} variance)")
    ax.set_ylabel(f"PC2 ({pca.explained_variance_ratio_[1]:.1%} variance)")
    ax.set_title("PCA Biplot: TB Delay Determinants")
    ax.grid(True, alpha=0.3)

    fig.tight_layout()
    output_file = FIGURES_DIR / "pca_biplot_delay_determinants.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved biplot to {output_file}")


def interpret_components(loadings: pd.DataFrame, explained_var: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    """Interpret the principal components."""
    interpretation = []

    # PC1: Usually the strongest component
    pc1_loadings = loadings["PC1"].abs().sort_values(ascending=False)
    pc1_top = pc1_loadings.head(3)
    pc1_desc = f"PC1 ({explained_var.loc[0, 'explained_variance_ratio']:.1%}): "
    pc1_desc += f"Strongest loadings from {', '.join(pc1_top.index)}"

    # PC2: Second component
    pc2_loadings = loadings["PC2"].abs().sort_values(ascending=False)
    pc2_top = pc2_loadings.head(3)
    pc2_desc = f"PC2 ({explained_var.loc[1, 'explained_variance_ratio']:.1%}): "
    pc2_desc += f"Strongest loadings from {', '.join(pc2_top.index)}"

    # PC3: Third component (if available)
    pc3_desc = ""
    if len(loadings.columns) > 2:
        pc3_loadings = loadings["PC3"].abs().sort_values(ascending=False)
        pc3_top = pc3_loadings.head(3)
        pc3_desc = f"PC3 ({explained_var.loc[2, 'explained_variance_ratio']:.1%}): "
        pc3_desc += f"Strongest loadings from {', '.join(pc3_top.index)}"

    interpretation_df = pd.DataFrame({
        "component": ["PC1", "PC2", "PC3"],
        "explained_variance_ratio": explained_var["explained_variance_ratio"].iloc[:3].values,
        "cumulative_variance": explained_var["cumulative_variance"].iloc[:3].values,
        "interpretation": [pc1_desc, pc2_desc, pc3_desc]
    })

    logger.info("Component interpretations:")
    for _, row in interpretation_df.iterrows():
        logger.info(f"  {row['component']}: {row['interpretation']}")

    return interpretation_df


def save_pca_results(state_data: pd.DataFrame, pca: PCA, loadings: pd.DataFrame,
                    explained_var: pd.DataFrame, interpretation: pd.DataFrame, logger: logging.Logger) -> None:
    """Save all PCA results to files."""
    # Component scores
    scores = pca.transform(StandardScaler().fit_transform(state_data))
    scores_df = pd.DataFrame(scores, columns=[f"PC{i+1}" for i in range(scores.shape[1])], index=state_data.index)
    scores_path = OUTPUT_DIR / "pca_component_scores_delay_determinants.csv"
    scores_df.to_csv(scores_path)
    logger.info(f"Saved component scores to {scores_path}")

    # Loadings
    loadings_path = OUTPUT_DIR / "pca_loadings_delay_determinants.csv"
    loadings.to_csv(loadings_path)
    logger.info(f"Saved loadings to {loadings_path}")

    # Explained variance
    explained_path = OUTPUT_DIR / "pca_explained_variance_delay_determinants.csv"
    explained_var.to_csv(explained_path, index=False)
    logger.info(f"Saved explained variance to {explained_path}")

    # Interpretation
    interp_path = OUTPUT_DIR / "pca_interpretation_delay_determinants.csv"
    interpretation.to_csv(interp_path, index=False)
    logger.info(f"Saved interpretation to {interp_path}")


def main() -> None:
    """Main execution function."""
    logger = configure_logging()
    logger.info("Starting PCA analysis of TB delay determinants")

    df = load_data(logger)
    if df.empty:
        logger.error("No data available for PCA analysis")
        return

    state_data, scaled_data = prepare_pca_data(df, logger)
    if state_data.empty:
        logger.error("Insufficient data for PCA")
        return

    pca, loadings, explained_var = perform_pca(scaled_data, state_data, logger)

    # Create visualizations
    create_scree_plot(explained_var, logger)
    create_loadings_heatmap(loadings, logger)
    create_component_biplot(pca, scaled_data, state_data, logger)

    # Interpret components
    interpretation = interpret_components(loadings, explained_var, logger)

    # Save results
    save_pca_results(state_data, pca, loadings, explained_var, interpretation, logger)

    logger.info("PCA analysis of TB delay determinants completed")


if __name__ == "__main__":
    main()
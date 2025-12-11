"""
Directed Acyclic Graph (DAG) causal analysis of TB detection delay determinants.

This script constructs and analyzes causal pathways between socioeconomic factors,
health system indicators, and TB detection delays using DAG methodology.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import pandas as pd

try:
    import networkx as nx
except ImportError:
    nx = None

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = PROJECT_ROOT / "data" / "processed" / "proxy_delay_results.csv"
OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"
FIGURES_DIR = PROJECT_ROOT / "output" / "figures"
LOG_PATH = OUTPUT_DIR / "dag_causal_delay_analysis.log"

# Define causal relationships with evidence strength
CAUSAL_RELATIONSHIPS = [
    # Socioeconomic → Risk factors
    ("poverty_pct", "symptomatic_no_care_pct", "strong"),
    ("literacy_pct", "private_first_provider_pct", "moderate"),
    ("literacy_pct", "bact_confirmed_pct", "moderate"),

    # Socioeconomic → System factors
    ("literacy_pct", "crowding_index", "weak"),
    ("poverty_pct", "crowding_index", "moderate"),

    # Risk factors → Delay
    ("symptomatic_no_care_pct", "pn_ratio", "strong"),
    ("private_first_provider_pct", "pn_ratio", "moderate"),

    # System factors → Delay
    ("bact_confirmed_pct", "pn_ratio", "strong"),
    ("crowding_index", "pn_ratio", "moderate"),

    # System factors → Incidence
    ("bact_confirmed_pct", "in_ratio", "moderate"),
    ("crowding_index", "in_ratio", "weak"),

    # Incidence → Delay
    ("in_ratio", "pn_ratio", "strong"),
]

# Node attributes for visualization
NODE_ATTRIBUTES = {
    "poverty_pct": {"category": "socioeconomic", "color": "lightcoral"},
    "literacy_pct": {"category": "socioeconomic", "color": "lightcoral"},
    "symptomatic_no_care_pct": {"category": "risk_factor", "color": "lightblue"},
    "private_first_provider_pct": {"category": "risk_factor", "color": "lightblue"},
    "bact_confirmed_pct": {"category": "system_factor", "color": "lightgreen"},
    "crowding_index": {"category": "system_factor", "color": "lightgreen"},
    "in_ratio": {"category": "outcome", "color": "gold"},
    "pn_ratio": {"category": "outcome", "color": "gold"},
}


def configure_logging() -> logging.Logger:
    """Configure logging for the script."""
    logger = logging.getLogger("dag_causal_delay_analysis")
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


def build_dag(logger: logging.Logger) -> nx.DiGraph:
    """Build the Directed Acyclic Graph for TB delay determinants."""
    if nx is None:
        logger.error("NetworkX required for DAG analysis")
        return nx.DiGraph()

    G = nx.DiGraph()

    # Add nodes with attributes
    for node, attrs in NODE_ATTRIBUTES.items():
        G.add_node(node, **attrs)

    # Add edges with evidence strength
    edge_colors = {"strong": "red", "moderate": "orange", "weak": "gray"}
    for source, target, strength in CAUSAL_RELATIONSHIPS:
        if source in G.nodes and target in G.nodes:
            G.add_edge(source, target, strength=strength, color=edge_colors[strength])

    # Verify DAG (no cycles)
    if not nx.is_directed_acyclic_graph(G):
        logger.warning("Graph contains cycles - not a valid DAG")
        return nx.DiGraph()

    logger.info(f"Built DAG with {len(G.nodes)} nodes and {len(G.edges)} edges")
    return G


def analyze_dag_structure(G: nx.DiGraph, logger: logging.Logger) -> Dict:
    """Analyze DAG structure and properties."""
    analysis = {
        "nodes": len(G.nodes),
        "edges": len(G.edges),
        "is_dag": nx.is_directed_acyclic_graph(G),
        "density": nx.density(G),
        "average_clustering": nx.average_clustering(G.to_undirected()),
    }

    # Node degrees
    in_degrees = dict(G.in_degree())
    out_degrees = dict(G.out_degree())

    analysis["max_in_degree"] = max(in_degrees.values()) if in_degrees else 0
    analysis["max_out_degree"] = max(out_degrees.values()) if out_degrees else 0

    # Pathways to outcomes
    outcomes = [node for node, attrs in G.nodes(data=True) if attrs.get("category") == "outcome"]
    pathways = {}

    for outcome in outcomes:
        predecessors = list(nx.ancestors(G, outcome))
        pathways[outcome] = {
            "direct_causes": list(G.predecessors(outcome)),
            "indirect_causes": [p for p in predecessors if p not in G.predecessors(outcome)],
            "total_pathways": len(list(nx.all_simple_paths(G, source=list(G.nodes())[0], target=outcome))) if G.nodes else 0
        }

    analysis["pathways"] = pathways

    logger.info(f"DAG analysis: {analysis['nodes']} nodes, {analysis['edges']} edges")
    return analysis


def visualize_dag(G: nx.DiGraph, logger: logging.Logger) -> None:
    """Create visualization of the DAG."""
    if not G.nodes:
        logger.warning("Empty graph - skipping visualization")
        return

    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    # Calculate layout
    pos = nx.spring_layout(G, k=2, iterations=50, seed=42)

    fig, ax = plt.subplots(figsize=(14, 10))

    # Draw nodes by category
    categories = {}
    for node, attrs in G.nodes(data=True):
        cat = attrs.get("category", "unknown")
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(node)

    colors = {"socioeconomic": "lightcoral", "risk_factor": "lightblue",
             "system_factor": "lightgreen", "outcome": "gold"}

    for cat, nodes in categories.items():
        nx.draw_networkx_nodes(G, pos, nodelist=nodes,
                              node_color=colors.get(cat, "gray"),
                              node_size=2000, alpha=0.8, ax=ax)

    # Draw edges by strength
    edge_colors = [G[u][v].get("color", "gray") for u, v in G.edges()]
    edge_widths = [{"strong": 3, "moderate": 2, "weak": 1}[G[u][v].get("strength", "weak")] for u, v in G.edges()]

    nx.draw_networkx_edges(G, pos, edge_color=edge_colors,
                          width=edge_widths, alpha=0.7, ax=ax,
                          arrows=True, arrowsize=20)

    # Draw labels
    labels = {node: node.replace("_pct", "%").replace("_", "\n") for node in G.nodes()}
    nx.draw_networkx_labels(G, pos, labels, font_size=8, font_weight="bold", ax=ax)

    # Add legend
    legend_elements = [
        plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='lightcoral', markersize=10, label='Socioeconomic'),
        plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='lightblue', markersize=10, label='Risk Factors'),
        plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='lightgreen', markersize=10, label='System Factors'),
        plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='gold', markersize=10, label='Outcomes'),
        plt.Line2D([0], [0], color='red', linewidth=3, label='Strong evidence'),
        plt.Line2D([0], [0], color='orange', linewidth=2, label='Moderate evidence'),
        plt.Line2D([0], [0], color='gray', linewidth=1, label='Weak evidence'),
    ]
    ax.legend(handles=legend_elements, loc='upper left', bbox_to_anchor=(1.05, 1))

    ax.set_title("Causal DAG: TB Detection Delay Determinants", fontsize=14, fontweight="bold")
    ax.axis('off')

    fig.tight_layout()
    output_file = FIGURES_DIR / "dag_causal_delay_analysis.png"
    fig.savefig(output_file, dpi=300, bbox_inches="tight")
    plt.close(fig)
    logger.info(f"Saved DAG visualization to {output_file}")


def analyze_causal_paths(G: nx.DiGraph, logger: logging.Logger) -> pd.DataFrame:
    """Analyze causal pathways and their importance."""
    if not G.nodes:
        return pd.DataFrame()

    paths_data = []

    # Find all simple paths to outcomes
    outcomes = [node for node, attrs in G.nodes(data=True) if attrs.get("category") == "outcome"]

    for outcome in outcomes:
        for source in G.nodes():
            if source == outcome:
                continue
            try:
                paths = list(nx.all_simple_paths(G, source, outcome))
                for path in paths:
                    if len(path) > 2:  # Only paths with intermediaries
                        path_length = len(path) - 1
                        evidence_strengths = []

                        # Calculate path evidence strength
                        for i in range(len(path) - 1):
                            edge_data = G.get_edge_data(path[i], path[i+1])
                            strength = edge_data.get("strength", "weak")
                            evidence_strengths.append(strength)

                        # Convert to numeric for averaging
                        strength_scores = {"strong": 3, "moderate": 2, "weak": 1}
                        avg_strength = np.mean([strength_scores[s] for s in evidence_strengths])

                        paths_data.append({
                            "outcome": outcome,
                            "source": source,
                            "path": " → ".join(path),
                            "path_length": path_length,
                            "evidence_strengths": ", ".join(evidence_strengths),
                            "avg_evidence_score": avg_strength,
                            "n_edges": len(evidence_strengths)
                        })

            except nx.NetworkXNoPath:
                continue

    paths_df = pd.DataFrame(paths_data)

    if not paths_df.empty:
        # Sort by evidence strength and path length
        paths_df = paths_df.sort_values(["outcome", "avg_evidence_score", "path_length"],
                                      ascending=[True, False, True])

    logger.info(f"Analyzed {len(paths_df)} causal pathways")
    return paths_df


def calculate_dag_metrics(G: nx.DiGraph, df: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    """Calculate DAG-based metrics for each state."""
    if df.empty or not G.nodes:
        return pd.DataFrame()

    # Aggregate state-level data
    state_data = df.groupby("state").mean(numeric_only=True)

    metrics = []

    for state in state_data.index:
        state_metrics = {"state": state}

        # Calculate node-level contributions
        for node in G.nodes():
            if node in state_data.columns:
                value = state_data.loc[state, node]

                # Calculate causal influence (sum of downstream paths)
                downstream = nx.descendants(G, node)
                influence_score = len(downstream)

                # Calculate evidence-weighted influence
                weighted_influence = 0
                for descendant in downstream:
                    try:
                        paths = list(nx.all_simple_paths(G, node, descendant))
                        for path in paths:
                            path_strength = 0
                            for i in range(len(path) - 1):
                                edge_data = G.get_edge_data(path[i], path[i+1])
                                strength = edge_data.get("strength", "weak")
                                strength_scores = {"strong": 3, "moderate": 2, "weak": 1}
                                path_strength += strength_scores[strength]
                            weighted_influence += path_strength / len(path)
                    except:
                        continue

                state_metrics[f"{node}_value"] = value
                state_metrics[f"{node}_influence"] = influence_score
                state_metrics[f"{node}_weighted_influence"] = weighted_influence

        metrics.append(state_metrics)

    metrics_df = pd.DataFrame(metrics)
    logger.info(f"Calculated DAG metrics for {len(metrics_df)} states")
    return metrics_df


def save_dag_results(G: nx.DiGraph, analysis: Dict, paths_df: pd.DataFrame,
                    metrics_df: pd.DataFrame, logger: logging.Logger) -> None:
    """Save all DAG analysis results."""
    # Graph structure
    edges_df = pd.DataFrame([
        {"source": u, "target": v, "strength": G[u][v].get("strength", "unknown")}
        for u, v in G.edges()
    ])
    edges_path = OUTPUT_DIR / "dag_edges_delay_analysis.csv"
    edges_df.to_csv(edges_path, index=False)
    logger.info(f"Saved DAG edges to {edges_path}")

    # Analysis summary
    analysis_path = OUTPUT_DIR / "dag_analysis_summary_delay.csv"
    pd.DataFrame([analysis]).to_csv(analysis_path, index=False)
    logger.info(f"Saved analysis summary to {analysis_path}")

    # Causal paths
    if not paths_df.empty:
        paths_path = OUTPUT_DIR / "dag_causal_paths_delay.csv"
        paths_df.to_csv(paths_path, index=False)
        logger.info(f"Saved causal paths to {paths_path}")

    # State metrics
    if not metrics_df.empty:
        metrics_path = OUTPUT_DIR / "dag_state_metrics_delay.csv"
        metrics_df.to_csv(metrics_path, index=False)
        logger.info(f"Saved state metrics to {metrics_path}")


def main() -> None:
    """Main execution function."""
    logger = configure_logging()
    logger.info("Starting DAG causal analysis of TB delay determinants")

    if nx is None:
        logger.error("NetworkX required for DAG analysis")
        return

    df = load_data(logger)
    if df.empty:
        logger.error("No data available for DAG analysis")
        return

    # Build DAG
    G = build_dag(logger)
    if not G.nodes:
        logger.error("Failed to build valid DAG")
        return

    # Analyze structure
    analysis = analyze_dag_structure(G, logger)

    # Visualize
    visualize_dag(G, logger)

    # Analyze pathways
    paths_df = analyze_causal_paths(G, logger)

    # Calculate metrics
    metrics_df = calculate_dag_metrics(G, df, logger)

    # Save results
    save_dag_results(G, analysis, paths_df, metrics_df, logger)

    logger.info("DAG causal analysis of TB delay determinants completed")


if __name__ == "__main__":
    main()
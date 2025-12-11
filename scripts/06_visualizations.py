"""Generate visualizations for TB delay analyses."""
from __future__ import annotations

import logging
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

try:
    import geopandas as gpd
    from shapely.geometry import Point
except ImportError:
    gpd = None
    Point = None

PROJECT_ROOT = Path(__file__).resolve().parents[1]
META_PATH = PROJECT_ROOT / "data" / "processed" / "meta_delay_results.csv"
PROXY_PATH = PROJECT_ROOT / "data" / "processed" / "proxy_delay_results.csv"
PANEL_PATH = PROJECT_ROOT / "data" / "processed" / "state_year_panel.csv"
FIG_DIR = PROJECT_ROOT / "output" / "figures"
LOG_PATH = PROJECT_ROOT / "output" / "figures" / "visualizations.log"
SHAPEFILE_PATH = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "shapefiles"
    / "ne_admin1"
    / "ne_50m_admin_1_states_provinces.shp"
)
STATE_COORDS = {
    "andaman and nicobar islands": (11.75, 92.72),
    "andhra pradesh": (15.91, 79.74),
    "arunachal pradesh": (28.21, 94.73),
    "assam": (26.20, 92.93),
    "bihar": (25.09, 85.31),
    "chandigarh": (30.74, 76.79),
    "chhattisgarh": (21.27, 82.04),
    "dadra and nagar haveli and daman and diu": (20.27, 72.90),
    "delhi": (28.61, 77.21),
    "goa": (15.49, 73.83),
    "gujarat": (22.25, 72.68),
    "haryana": (29.06, 76.09),
    "himachal pradesh": (31.10, 77.17),
    "jammu and kashmir": (34.08, 76.83),
    "jharkhand": (23.61, 85.28),
    "karnataka": (15.32, 75.76),
    "kerala": (10.35, 76.27),
    "ladakh": (34.15, 77.58),
    "lakshadweep": (10.57, 72.64),
    "madhya pradesh": (23.52, 80.83),
    "maharashtra": (19.76, 75.71),
    "manipur": (24.72, 93.91),
    "meghalaya": (25.47, 91.37),
    "mizoram": (23.46, 92.69),
    "nagaland": (26.16, 94.56),
    "odisha": (20.95, 85.10),
    "puducherry": (11.91, 79.81),
    "punjab": (31.15, 75.34),
    "rajasthan": (26.91, 74.80),
    "sikkim": (27.53, 88.52),
    "tamil nadu": (11.12, 78.66),
    "telangana": (17.87, 79.60),
    "tripura": (23.84, 91.29),
    "uttar pradesh": (26.85, 80.91),
    "uttarakhand": (30.18, 79.30),
    "west bengal": (23.16, 87.84),
    "andaman & nicobar islands": (11.75, 92.72),
    "dadra and nagar haveli": (20.27, 72.90),
    "daman and diu": (20.41, 72.84),
    "maharastra": (19.76, 75.71),
}

STATE_NAME_OVERRIDES = {
    "dadra and nagar haveli": "dadra and nagar haveli and daman and diu",
    "daman and diu": "dadra and nagar haveli and daman and diu",
    "nct of delhi": "delhi",
    "maharastra": "maharashtra",
}

sns.set_style("whitegrid")


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("visualizations")
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


def safe_read(path: Path) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame()
    return pd.read_csv(path)


def save_placeholder(filename: str, message: str) -> None:
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.axis("off")
    ax.text(0.5, 0.5, message, ha="center", va="center", wrap=True)
    fig.savefig(FIG_DIR / filename, dpi=200)
    plt.close(fig)


def normalize_state_name(name: str) -> str:
    if not isinstance(name, str):
        return ""
    lowered = name.lower().replace("&", " and ")
    cleaned = "".join(ch if ch.isalnum() or ch.isspace() else " " for ch in lowered)
    normalized = " ".join(cleaned.split())
    return STATE_NAME_OVERRIDES.get(normalized, normalized)


def aggregate_state_metrics(proxies: pd.DataFrame) -> pd.DataFrame:
    if proxies.empty or "state" not in proxies.columns:
        return pd.DataFrame()
    agg = (
        proxies.groupby("state")
        .agg(pn_ratio=("pn_ratio", "mean"), delay_cluster=("delay_cluster", "first"))
        .reset_index()
    )
    agg["state_norm"] = agg["state"].apply(normalize_state_name)
    return agg


def plot_forest_from_meta(meta_df: pd.DataFrame, logger: logging.Logger) -> None:
    if meta_df.empty:
        save_placeholder("forest_plot.png", "Meta-analysis results unavailable")
        logger.warning("Meta-analysis results missing; placeholder forest plot created.")
        return
    fig, ax = plt.subplots(figsize=(6, 4 + len(meta_df) * 0.5))
    ax.errorbar(
        meta_df["effect"],
        range(len(meta_df)),
        xerr=1.96 * meta_df["se"],
        fmt="o",
        color="darkblue",
    )
    ax.set_yticks(range(len(meta_df)))
    ax.set_yticklabels(meta_df["delay_type"])
    ax.set_xlabel("Delay (days)")
    ax.set_title("Pooled TB delay estimates")
    fig.tight_layout()
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(FIG_DIR / "forest_plot.png", dpi=300)
    plt.close(fig)
    logger.info("Forest plot saved to output/figures/forest_plot.png")


def plot_state_heatmap(proxies: pd.DataFrame, logger: logging.Logger) -> None:
    if proxies.empty or "year" not in proxies.columns or "pn_ratio" not in proxies.columns:
        save_placeholder("state_heatmap.png", "State-year proxy data missing")
        logger.warning("Proxy dataset missing year or P:N ratio column; cannot draw heatmap.")
        return
    heatmap_data = proxies.pivot_table(
        index="state", columns="year", values="pn_ratio", aggfunc="mean"
    )
    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(heatmap_data, cmap="viridis", ax=ax, cbar_kws={"label": "P:N ratio"})
    ax.set_title("State P:N ratio heatmap")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "state_heatmap.png", dpi=300)
    plt.close(fig)
    logger.info("State heatmap saved.")


def plot_cluster_map(proxies: pd.DataFrame, logger: logging.Logger) -> None:
    if "delay_cluster" not in proxies.columns:
        save_placeholder("cluster_map.png", "Cluster assignments unavailable")
        logger.warning("Delay cluster column missing.")
        return
    cluster_counts = proxies.groupby(["state", "delay_cluster"]).size().reset_index(name="count")
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.barplot(
        data=cluster_counts,
        x="state",
        y="count",
        hue="delay_cluster",
        ax=ax,
    )
    ax.set_title("Cluster membership counts by state")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "cluster_map.png", dpi=300)
    plt.close(fig)
    logger.info("Cluster map saved.")


def plot_scatter(proxies: pd.DataFrame, logger: logging.Logger) -> None:
    required = {"crowding_index", "pn_ratio"}
    if proxies.empty or not required.issubset(set(proxies.columns)):
        save_placeholder("scatter_proxy.png", "No proxy data available")
        logger.warning("Cannot build scatter plot; required columns missing.")
        return
    fig, ax = plt.subplots(figsize=(6, 4))
    sns.scatterplot(
        data=proxies,
        x="crowding_index",
        y="pn_ratio",
        hue="delay_cluster",
        ax=ax,
    )
    ax.set_title("Crowding vs P:N ratio")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "scatter_proxy.png", dpi=300)
    plt.close(fig)
    logger.info("Scatter plot saved.")


def plot_radar_charts(proxies: pd.DataFrame, logger: logging.Logger) -> None:
    features = ["pn_ratio", "in_ratio", "crowding_index", "literacy_pct", "poverty_pct"]
    if (
        proxies.empty
        or "state" not in proxies.columns
        or not set(features).issubset(proxies.columns)
    ):
        save_placeholder("radar_placeholder.png", "Insufficient features for radar charts")
        logger.warning("Cannot render radar charts; missing features.")
        return
    for state, group in proxies.groupby("state"):
        values = group[features].mean().fillna(0).tolist()
        angles = np.linspace(0, 2 * np.pi, len(features), endpoint=False).tolist()
        values += values[:1]
        angles += angles[:1]
        fig, ax = plt.subplots(figsize=(5, 5), subplot_kw={"polar": True})
        ax.plot(angles, values, label=state)
        ax.fill(angles, values, alpha=0.2)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(features, fontsize=9)
        ax.set_title(f"Timeliness profile: {state}")
        fig.tight_layout()
        safe_state = state.lower().replace(" ", "_").replace("/", "_")
        filename = FIG_DIR / f"radar_{safe_state}.png"
        fig.savefig(filename, dpi=250)
        plt.close(fig)
    logger.info("Radar charts saved for %s states.", proxies['state'].nunique())


def plot_geopandas_map(proxies: pd.DataFrame, logger: logging.Logger) -> None:
    if gpd is None or Point is None:
        save_placeholder("state_geopandas_map.png", "GeoPandas not installed")
        logger.warning("GeoPandas not available; skipping geospatial map.")
        return
    aggregated = aggregate_state_metrics(proxies)
    if aggregated.empty:
        save_placeholder("state_geopandas_map.png", "Proxy dataset missing for GeoPandas map")
        logger.warning("Insufficient proxy data for GeoPandas map.")
        return
    records = []
    missing_states = []
    for _, row in aggregated.iterrows():
        norm = row["state_norm"]
        coords = STATE_COORDS.get(norm)
        if not coords:
            missing_states.append(row["state"])
            continue
        lat, lon = coords
        records.append(
            {
                "state": row["state"],
                "pn_ratio": row["pn_ratio"],
                "geometry": Point(lon, lat),
            }
        )
    if not records:
        save_placeholder("state_geopandas_map.png", "No state coordinates available")
        logger.warning("No matching coordinates found for any state.")
        return
    geo_df = gpd.GeoDataFrame(records, crs="EPSG:4326")
    fig, ax = plt.subplots(figsize=(7, 8))
    geo_df.plot(column="pn_ratio", cmap="OrRd", legend=True, markersize=60, ax=ax)
    ax.set_title("State-level TB prevalence-to-notification proxy (centroid map)")
    ax.set_xlabel("Longitude")
    ax.set_ylabel("Latitude")
    fig.tight_layout()
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(FIG_DIR / "state_geopandas_map.png", dpi=300)
    plt.close(fig)
    if missing_states:
        logger.warning(
            "Missing coordinates for %s states: %s",
            len(missing_states),
            ", ".join(sorted(missing_states)),
        )


def plot_shapefile_map(proxies: pd.DataFrame, logger: logging.Logger) -> None:
    if gpd is None:
        save_placeholder("state_shapefile_map.png", "GeoPandas not installed")
        logger.warning("GeoPandas not available; cannot plot shapefile map.")
        return
    if not SHAPEFILE_PATH.exists():
        save_placeholder("state_shapefile_map.png", "Shapefile not found")
        logger.warning("Shapefile %s missing.", SHAPEFILE_PATH)
        return
    aggregated = aggregate_state_metrics(proxies)
    if aggregated.empty:
        save_placeholder("state_shapefile_map.png", "No proxy data for shapefile map")
        logger.warning("Proxy data unavailable for shapefile join.")
        return
    try:
        world_admin = gpd.read_file(SHAPEFILE_PATH)
    except Exception as exc:  # pragma: no cover
        save_placeholder("state_shapefile_map.png", f"Failed to load shapefile: {exc}")
        logger.warning("Unable to load shapefile: %s", exc)
        return
    india = world_admin[world_admin["adm0_a3"] == "IND"].copy()
    if india.empty:
        save_placeholder("state_shapefile_map.png", "No India polygons in shapefile")
        logger.warning("Shapefile does not contain India polygons.")
        return
    india["state_norm"] = india["name"].apply(normalize_state_name)
    merged = india.merge(aggregated, on="state_norm", how="left")
    fig, ax = plt.subplots(figsize=(8, 10))
    merged.plot(
        column="pn_ratio",
        cmap="YlOrRd",
        linewidth=0.5,
        edgecolor="gray",
        missing_kwds={"color": "lightgray", "label": "No data"},
        legend=True,
        ax=ax,
    )
    ax.set_title("India state-level P:N ratio (shapefile choropleth)")
    ax.axis("off")
    fig.tight_layout()
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(FIG_DIR / "state_shapefile_map.png", dpi=300)
    plt.close(fig)


def main() -> None:
    logger = configure_logging()
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    meta_df = safe_read(META_PATH)
    proxy_df = safe_read(PROXY_PATH)
    plot_forest_from_meta(meta_df, logger)
    plot_state_heatmap(proxy_df, logger)
    plot_cluster_map(proxy_df, logger)
    plot_scatter(proxy_df, logger)
    plot_radar_charts(proxy_df, logger)
    plot_geopandas_map(proxy_df, logger)
    plot_shapefile_map(proxy_df, logger)


if __name__ == "__main__":
    main()

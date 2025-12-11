"""Advanced visualizations for TB delay analysis with GIS, geopandas, and innovative charts."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap

try:
    import geopandas as gpd
    from shapely.geometry import Point
    GEOSPATIAL_AVAILABLE = True
except ImportError:
    GEOSPATIAL_AVAILABLE = False

PROJECT_ROOT = Path(__file__).resolve().parents[1]
PROXY_PATH = PROJECT_ROOT / "data" / "processed" / "proxy_delay_results.csv"
META_PATH = PROJECT_ROOT / "data" / "processed" / "meta_delay_results.csv"
FIGURES_PATH = PROJECT_ROOT / "output" / "figures"
TABLES_PATH = PROJECT_ROOT / "output" / "tables"

# Publication-ready styling
plt.rcParams.update({
    'font.size': 11,
    'font.family': 'serif',
    'font.serif': ['Times New Roman', 'DejaVu Serif'],
    'axes.labelsize': 12,
    'axes.titlesize': 13,
    'figure.titlesize': 14,
    'legend.fontsize': 10,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'figure.figsize': (10, 6),
    'figure.dpi': 300,
    'savefig.dpi': 300,
    'savefig.bbox': 'tight'
})

# TB color scheme
TB_COLORS = {
    'primary': '#2E86AB',      # Teal blue (official TB color)
    'delay_high': '#F24236',   # Red for high delay
    'delay_med': '#FFB400',    # Orange/yellow for medium
    'delay_low': '#00A651',    # Green for low delay
    'background': '#F8F9FA',
    'cluster_0': '#00A651',    # Best performing
    'cluster_1': '#FFB400',    # Moderate
    'cluster_2': '#F24236',    # High risk
    'cluster_3': '#8B4513',    # Metropolitan challenge
}

def load_data():
    """Load all necessary datasets for visualization."""
    proxy_df = pd.read_csv(PROXY_PATH)
    meta_df = pd.read_csv(META_PATH)

    return proxy_df, meta_df

def create_geospatial_choropleth(proxy_df):
    """Create geospatial choropleth map using geopandas."""
    if not GEOSPATIAL_AVAILABLE:
        print("GeoPandas not available, skipping geospatial visualizations")
        return

    try:
        # Load shapefiles
        shapefile_path = PROJECT_ROOT / "data" / "raw" / "shapefiles"
        if not shapefile_path.exists():
            print("Shapefiles not found, skipping geospatial maps")
            return

        # Create bubble map with approximate India boundaries
        fig, ax = plt.subplots(figsize=(15, 10))

        # States data with approximate coordinates
        india_states_coords = {
            'Maharashtra': [72.5, 19.4], 'Uttar Pradesh': [80.3, 25.8],
            'Bihar': [85.7, 25.6], 'West Bengal': [87.8, 22.6],
            'Madhya Pradesh': [78.6, 22.7], 'Rajasthan': [74.2, 26.9],
            'Gujarat': [71.3, 22.3], 'Andhra Pradesh': [79.7, 15.9],
            'Karnataka': [75.7, 15.3], 'Tamil Nadu': [78.7, 11.1],
            'Telangana': [79.0, 18.1], 'Punjab': [75.3, 31.1]
        }

        gdf_data = []
        for state, coords in india_states_coords.items():
            state_data = proxy_df[proxy_df['state'] == state]
            if not state_data.empty:
                gdf_data.append({
                    'state': state,
                    'lon': coords[0],
                    'lat': coords[1],
                    'delay_cluster': state_data['delay_cluster'].iloc[0],
                    'pn_ratio': state_data['pn_ratio'].iloc[0]
                })

        # Convert to DataFrame
        bubble_df = pd.DataFrame(gdf_data)

        # Create bubble map
        cluster_colors = {0: TB_COLORS['cluster_0'], 1: TB_COLORS['cluster_1'],
                         2: TB_COLORS['cluster_2'], 3: TB_COLORS['cluster_3']}

        for _, row in bubble_df.iterrows():
            cluster_color = cluster_colors.get(row['delay_cluster'], TB_COLORS['primary'])
            size = max(100, min(1000, row['pn_ratio'] * 5000)) if pd.notna(row['pn_ratio']) else 300

            ax.scatter(row['lon'], row['lat'], s=size, c=cluster_color,
                      alpha=0.7, edgecolors='black', linewidth=1)

            # Add state labels
            ax.annotate(row['state'], (row['lon'], row['lat']),
                       fontsize=9, ha='center', va='bottom', fontweight='bold')

        # Approximate India boundaries
        ax.set_xlim([65, 95])
        ax.set_ylim([10, 35])

        ax.set_title('Geospatial Bubble Map: TB Delay Intensity Across Indian States\n' +
                    'Bubble Size: Prevalence-to-Notification Ratio | Color: Delay Cluster',
                    fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Longitude')
        ax.set_ylabel('Latitude')
        ax.grid(True, alpha=0.3)

        plt.tight_layout()
        fig.savefig(FIGURES_PATH / 'gis_bubble_map.png',
                    dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        print("Geospatial bubble map created successfully")

    except Exception as e:
        print(f"Error creating geospatial visualization: {e}")

def create_innovative_data_storytelling_plot():
    """Create innovative data storytelling visualization."""
    proxy_df, meta_df = load_data()

    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))

    # Panel 1: Circular bar chart of state delays
    states_top = proxy_df.nlargest(15, 'pn_ratio')
    angles = np.linspace(0, 2 * np.pi, len(states_top), endpoint=False)

    ax1.bar(angles, states_top['pn_ratio'], width=0.4, color=TB_COLORS['delay_high'], alpha=0.7)
    ax1.set_xticks(angles)
    ax1.set_xticklabels(states_top['state'], rotation=45, ha='right', fontsize=8)
    ax1.set_title('Top 15 States: Delay Intensity\n(Circular Bar Chart)', fontweight='bold')
    ax1.grid(True, alpha=0.3)

    # Panel 2: Dumbbell plot for cluster comparison
    cluster_means = proxy_df.groupby('delay_cluster')[['pn_ratio', 'symptomatic_no_care_pct']].mean()

    for i, (_, row) in enumerate(cluster_means.iterrows()):
        cluster_color = [TB_COLORS['cluster_0'], TB_COLORS['cluster_1'],
                        TB_COLORS['cluster_2'], TB_COLORS['cluster_3']][i]

        ax2.plot([row['pn_ratio'], row['symptomatic_no_care_pct']],
                [i, i], 'o-', linewidth=3, markersize=10, color=cluster_color)

    cluster_labels = [f'Cluster {i}' for i in range(len(cluster_means))]
    ax2.set_yticks(range(len(cluster_means)))
    ax2.set_yticklabels(cluster_labels, fontsize=10, fontweight='bold')
    ax2.set_xlabel('Scale')
    ax2.set_title('Cluster Comparison: Epidemic Intelligence\n(Dumbbell Plot)', fontweight='bold')
    ax2.legend(['P:N Ratio', 'Symptomatic Non-Care'], loc='upper right', fontsize=9)
    ax2.grid(True, alpha=0.3)

    # Panel 3: Waterfall chart of delay components
    delay_components = ['Patient\nDelay', 'Diagnostic\nDelay', 'Treatment\nDelay']
    values = [18.43, 29.40, 4.0]
    cumulative = [0] + list(np.cumsum(values)[:-1])

    for i, (component, value, cum) in enumerate(zip(delay_components, values, cumulative)):
        color = [TB_COLORS['primary'], TB_COLORS['delay_high'], TB_COLORS['delay_med']][i]
        ax3.bar(i, value, bottom=cum, color=color, alpha=0.8, edgecolor='black', linewidth=1)

    ax3.bar(3, 49.17 - sum(values), bottom=sum(values),
           color=TB_COLORS['delay_low'], alpha=0.8, edgecolor='black', linewidth=1, label='Estimated')

    ax3.set_xticks(range(4))
    ax3.set_xticklabels(delay_components + ['Total\nDelay'])
    ax3.set_ylabel('Days')
    ax3.set_title('Delay Decomposition: Patient vs. System\n(Waterfall Chart)', fontweight='bold')
    ax3.grid(True, alpha=0.3)

    # Panel 4: Heatmap of state clusters
    cluster_pivot = proxy_df.pivot_table(
        values='pn_ratio',
        index='delay_cluster',
        columns='state',
        aggfunc='mean'
    ).fillna(0)

    top_states_by_cluster = []
    for cluster in range(4):
        cluster_states = proxy_df[proxy_df['delay_cluster'] == cluster].nlargest(5, 'pn_ratio')['state']
        top_states_by_cluster.extend(cluster_states)

    cluster_matrix = proxy_df[proxy_df['state'].isin(top_states_by_cluster[:10])].pivot_table(
        values='pn_ratio',
        index='delay_cluster',
        columns='state',
        aggfunc='mean'
    ).fillna(proxy_df['pn_ratio'].mean())

    sns.heatmap(cluster_matrix, ax=ax4, cmap='RdYlBu_r',
                square=True, cbar_kws={'shrink': 0.6})
    ax4.set_title('State-Cluster Delay Matrix\n(Heatmap)', fontweight='bold')

    plt.suptitle('Innovative TB Delay Analysis: Multi-Perspective Data Visualization\n' +
                'Integration of Policy Intelligence and Technical Excellence',
                fontsize=18, fontweight='bold', y=0.98)

    plt.tight_layout()
    fig.savefig(FIGURES_PATH / 'innovative_multipanel_dashboard.png',
                dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()

    return fig

def create_publication_tables():
    """Create publication-quality tables."""
    TABLES_PATH.mkdir(exist_ok=True)

    proxy_df, meta_df = load_data()

    # Table 1: Meta-analysis results
    meta_table_1 = meta_df.copy()
    meta_table_1['ci'] = meta_table_1.apply(lambda x: f"{x['ci_low']:.2f}-{x['ci_high']:.2f}", axis=1)
    meta_table_1['effect_ci'] = meta_table_1.apply(lambda x: f"{x['effect']:.2f} ({x['ci_low']:.2f}-{x['ci_high']:.2f})", axis=1)
    meta_table_1 = meta_table_1[['delay_type', 'effect_ci', 'k']]
    meta_table_1.columns = ['Delay Component', 'Mean (95% CI)', 'Studies']
    meta_table_1.to_csv(TABLES_PATH / 'table_1_meta_analysis.csv', index=False)

    # Table 2: State-level results top 10
    state_table = proxy_df.nlargest(10, 'pn_ratio')[['state', 'pn_ratio', 'delay_cluster',
                                                   'symptomatic_no_care_pct']].copy()
    state_table.columns = ['State', 'P:N Ratio', 'Cluster', 'Non-Care Rate (%)']
    state_table['P:N Ratio'] = state_table['P:N Ratio'].round(4)
    state_table['Non-Care Rate (%)'] = state_table['Non-Care Rate (%)'].round(1)
    state_table.to_csv(TABLES_PATH / 'table_2_state_rankings.csv', index=False)

    # Table 3: Cluster summary statistics
    cluster_stats = proxy_df.groupby('delay_cluster').agg({
        'pn_ratio': ['count', 'mean', 'std'],
        'symptomatic_no_care_pct': ['mean'],
        'private_first_provider_pct': ['mean'],
        'poverty_pct': ['mean']
    }).round(2)

    cluster_stats.columns = ['States', 'P:N Mean', 'P:N STD', 'Non-Care (%)',
                           'Private Contact (%)', 'Poverty (%)']
    cluster_stats.to_csv(TABLES_PATH / 'table_3_cluster_summary.csv')

def main():
    """Main execution function."""
    print("=== Creating Advanced Visualizations for TB Delay Analysis ===")

    # Create directories
    FIGURES_PATH.mkdir(exist_ok=True)
    TABLES_PATH.mkdir(exist_ok=True)

    # Load data
    proxy_df, meta_df = load_data()
    print(f"Loaded data: {len(proxy_df)} states, {len(meta_df)} delay types")

    # Create advanced visualizations
    print("Creating geospatial bubble map...")
    create_geospatial_choropleth(proxy_df)

    print("Creating innovative multi-panel dashboard...")
    create_innovative_data_storytelling_plot()

    print("Generating publication-quality tables...")
    create_publication_tables()

    print("=== Advanced Visualization Suite Complete ===")
    print("All visualizations saved to:", FIGURES_PATH)
    print("All tables saved to:", TABLES_PATH)

if __name__ == "__main__":
    main()

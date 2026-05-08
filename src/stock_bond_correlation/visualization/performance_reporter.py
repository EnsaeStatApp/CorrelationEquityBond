import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from typing import Dict, List

def plot_equity_curves(results: Dict, strategies: List[str] = None):
    """
    Plots equity curves with a 2-year frequency on the X-axis and vertical grid lines.

    Args:
        results (Dict): Dictionary containing 'dates' and return series.
        strategies (List[str], optional): List of keys from results to plot.
    """
    dates = results["dates"]
    
    # Mapping between internal keys and display labels
    label_mapping = {
        "portfolio_rets": "HMM TVTP (With Overlays)",
        "rets_hmm_pure": "HMM TVTP (100% Invested)",
        "rets_6040": "Benchmark 60/40",
        "rets_rp": "Rolling Risk Parity"
    }
    
    # Visual style configuration
    style_mapping = {
        "portfolio_rets": {"color": "black", "lw": 2.5, "alpha": 1.0},
        "rets_hmm_pure": {"color": "blue", "lw": 1.5, "alpha": 0.7},
        "rets_6040": {"color": "gray", "lw": 1.5, "alpha": 0.5},
        "rets_rp": {"color": "red", "lw": 1.5, "alpha": 0.6, "ls": "--"}
    }

    if strategies is None:
        strategies = [k for k in label_mapping.keys() if k in results]

    plt.figure(figsize=(15, 8))
    ax = plt.gca()
    
    for key in strategies:
        if key in results:
            # Calculate Wealth Index (Base 100)
            wealth = np.exp(np.cumsum(results[key])) * 100
            style = style_mapping.get(key, {"lw": 1.5})
            plt.plot(dates, wealth, label=label_mapping.get(key, key), **style)

    # --- X-AXIS FORMATTING (2-YEAR STEPS) ---
    ax.xaxis.set_major_locator(mdates.YearLocator(2))  # Tick every 2 years
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    
    # --- GRID & LABELS ---
    plt.title("Net Comparative Performance (Linear Scale)", fontsize=16, fontweight='bold')
    plt.ylabel("Wealth Index (Base 100)", fontsize=13)
    plt.xlabel("Date", fontsize=13)
    
    # Major grid for the 2-year ticks (Vertical lines)
    plt.grid(True, which='major', axis='x', color='gray', linestyle='--', alpha=0.4)
    # Horizontal grid for wealth levels
    plt.grid(True, which='major', axis='y', color='gray', linestyle='-', alpha=0.1)
    
    plt.tick_params(axis='both', which='major', labelsize=13)
    plt.legend(fontsize=12, loc='upper left', frameon=True)
    
    plt.tight_layout()
    plt.show()


def plot_allocation_stack(results: Dict):
    """
    Plots the stackplot of asset allocations over time.
    """
    dates = results["dates"]
    w_df = pd.DataFrame(
        results["history_weights"],
        index=dates,
        columns=["S&P 500", "T-Bond", "Cash"]
    )
    
    plt.figure(figsize=(15, 7))
    ax = plt.gca()
    
    plt.stackplot(
        dates,
        w_df["S&P 500"],
        w_df["T-Bond"],
        w_df["Cash"],
        labels=["S&P 500", "Bonds (T-Bond)", "Cash"],
        colors=['#1f777b', '#ff7f0e', '#2ca02c'],
        alpha=0.8
    )
    
    # Same X-axis logic for the stackplot to keep charts aligned
    ax.xaxis.set_major_locator(mdates.YearLocator(2))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    
    plt.title("HMM TVTP Dynamic Allocation Decomposition", fontsize=16, fontweight='bold')
    plt.ylabel("Portfolio Weights", fontsize=13)
    plt.ylim(0, 1)
    
    plt.tick_params(axis='both', which='major', labelsize=13)
    plt.legend(loc='upper left', fontsize=14, frameon=True, bbox_to_anchor=(1, 1))
    
    # Vertical grid for better readability
    plt.grid(True, which='major', axis='x', color='black', linestyle='--', alpha=0.1)
    
    plt.tight_layout()
    plt.show()

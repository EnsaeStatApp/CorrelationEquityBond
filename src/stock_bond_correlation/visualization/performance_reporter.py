import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from typing import Dict

def plot_equity_curves(results: Dict):
    """
    Plots the equity curves from simulation results using a linear scale.
    
    A linear scale is preferred here to better 
    visualize absolute drawdowns and performance dynamics in the post-2009 period,
    especially when total growth is around 4x.
    """
    dates = results["dates"]
    plt.figure(figsize=(15, 8))
    
    # Calculate Wealth Index (Base 100)
    strategy_wealth = np.exp(np.cumsum(results["portfolio_rets"])) * 100
    benchmark_wealth = np.exp(np.cumsum(results["rets_6040"])) * 100
    
    plt.plot(dates, strategy_wealth, label="HMM TVTP (With Overlays)", lw=2.5, color='black')
    plt.plot(dates, benchmark_wealth, label="Benchmark 60/40", alpha=0.5)
    
    # Removed plt.yscale("log") to switch to linear scale
    plt.title("Net Comparative Performance (Linear Scale)", fontsize=16, fontweight='bold')
    plt.ylabel("Wealth Index (Base 100)", fontsize=13)
    
    # Adjust tick sizes for readability
    plt.tick_params(axis='both', which='major', labelsize=13)
    
    plt.legend(fontsize=14, loc='upper left', frameon=True)
    plt.grid(True, alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    plt.show()

def plot_allocation_stack(results: Dict):
    """
    Plots the stackplot of asset allocations over time to visualize 
    dynamic exposures (Equity, Bonds, and Cash/Deleveraging).
    """
    dates = results["dates"]
    # history_weights contains [W_SP500, W_TBOND, W_CASH]
    w_df = pd.DataFrame(results["history_weights"], index=dates, columns=["S&P 500", "T-Bond", "Cash"])
    
    plt.figure(figsize=(15, 7))
    plt.stackplot(dates, w_df["S&P 500"], w_df["T-Bond"], w_df["Cash"],
                  labels=["S&P 500", "T-Bond", "Cash"], 
                  colors=['#1f777b', '#ff7f0e', '#2ca02c'], alpha=0.8)
    
    plt.title("HMM TVTP Dynamic Allocation Decomposition", fontsize=16, fontweight='bold')
    plt.ylabel("Portfolio Weights", fontsize=13)
    plt.ylim(0, 1) # Normalizing the view to 100%
    
    plt.tick_params(axis='both', which='major', labelsize=13)
    plt.legend(loc='upper left', fontsize=14, frameon=True)
    plt.grid(axis='y', alpha=0.2)
    
    plt.tight_layout()
    plt.show()
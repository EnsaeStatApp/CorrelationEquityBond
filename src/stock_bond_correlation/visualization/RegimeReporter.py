import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import Patch
import numpy as np
import pandas as pd
from typing import Union, List, Dict, Tuple
import matplotlib.dates as mdates


class RegimeReporter:
    """
    Class dedicated to the temporal visualization of detected regimes.
    
    Provides tools to overlay regime backgrounds on financial series, 
    annotate historical events, and visualize time-varying risk metrics.
    """
    def __init__(self, detector, log_returns: np.ndarray, dates: Union[list, np.ndarray, pd.DatetimeIndex],
                 states: np.ndarray, asset_names: List[str] = None,
                 regime_labels: Dict[int, str] = None, regime_colors: List[str] = None,
                 events_dict: Dict[str, str] = None, zoom_mode: bool = False):
        """
        Initializes the Regime Reporter.

        Args:
            detector: Fitted regime detector (Market or Macro).
            log_returns (np.ndarray): Array containing asset log returns.
            dates (Union): Time index corresponding to log_returns and states.
            states (np.ndarray): Sequence of detected regimes of size (T,). 
                                 Represents the dominant state (Viterbi or Filtered).
            asset_names (List[str], optional): Names of the D financial assets.
            regime_labels (Dict[int, str], optional): Mapping from regime ID to descriptive name.
            regime_colors (List[str], optional): Color codes for each regime.
            events_dict (Dict[str, str], optional): Historical events to annotate (e.g., {"COVID": "2020-03-01"}).
            zoom_mode (bool): If True, applies high-resolution formatting to the x-axis.
        """
        self.detector = detector
        self.log_returns = log_returns
        self.dates = pd.to_datetime(dates)
        self.states = states
        self.zoom_mode = zoom_mode
        self.K = detector.K
        self.D = log_returns.shape[1]

        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regime_labels = regime_labels or {k: f"Regime {k}" for k in range(self.K)}

        # Harmonized colors
        default_colors = ["#3498db", "#e74c3c", "#2ecc71", "#f1c40f", "#9b59b6"]
        self.regime_colors = {k: default_colors[k % len(default_colors)] for k in range(self.K)} if regime_colors is None else {k: regime_colors[k] for k in range(min(len(regime_colors), self.K))}

        self.events_dict = events_dict or {}

    def _shade_regimes(self, ax):
        """
        Adds colored background spans according to the state sequence.
        
        Spans extend from the first to the last index of a regime period,
        with a minimum one-month extension for single-point regimes.
        """
        n = len(self.dates)
        changes = np.where(self.states[:-1] != self.states[1:])[0]
        splits = np.concatenate(([0], changes + 1, [n]))
    
        for i in range(len(splits) - 1):
            start_idx = int(splits[i])
            end_idx = int(splits[i + 1] - 1)
            regime = self.states[start_idx]
    
            x0 = self.dates[start_idx]
    
            # Right boundary: next date if available, else +31 days
            if end_idx + 1 < n:
                x1 = self.dates[end_idx + 1]
            else:
                x1 = self.dates[end_idx] + pd.Timedelta(days=31)
    
            ax.axvspan(x0, x1,
                       alpha=0.25,
                       color=self.regime_colors.get(regime, "gray"),
                       linewidth=0,
                       zorder=0)
    
            ax.axvspan(x0, x1,
                       alpha=0.25,
                       color=self.regime_colors.get(regime, "gray"),
                       linewidth=0,
                       zorder=0)
            
    def _regime_patches(self):
        """
        Generates legend labels (patches) dynamically according to K.
        """
        return [Patch(facecolor=color, alpha=0.5,
                      label=self.regime_labels.get(k, f"Regime {k}"))
                for k, color in self.regime_colors.items() if k < self.K]

    def _annotate_events(self, ax):
        """
        Adds vertical dotted lines and labels for provided historical events.
        """
        if not self.events_dict:
            return

        for label, d in self.events_dict.items():
            d = pd.Timestamp(d)
            if self.dates.min() <= d <= self.dates.max():
                ax.axvline(d, color="#2c3e50", linewidth=0.8, alpha=0.4, linestyle=":")
                ymin, ymax = ax.get_ylim()
                
                # Correct positioning for log scale axes
                if ax.get_yscale() == "log":
                    log_ypos = np.log10(ymin) + 0.92 * (np.log10(ymax) - np.log10(ymin))
                    ypos = 10 ** log_ypos
                else:
                    ypos = ymin + (ymax - ymin) * 0.92
                ax.annotate(label, xy=(d, ypos), fontsize=8, fontweight="bold",
                            ha="center", va="top", color="#2c3e50", alpha=0.8)

    def _format_xaxis(self, ax):
        """
        Fixed formatting for monthly data:
        - Major: Every Year (Large tick + Year label) or 5 Years
        - Minor: Every Month (Small tick)
        - Vertical Grid: Major lines only
        """
        if self.zoom_mode:
            # 1. Ticks positioning (Locators)
            ax.xaxis.set_major_locator(mdates.YearLocator())
            ax.xaxis.set_minor_locator(mdates.MonthLocator())

            # 2. Text Format (Formatter)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))

            # 3. Tick Style
            ax.tick_params(axis='x', which='major', length=10, width=1.5, labelsize=10) # Major for years
            ax.tick_params(axis='x', which='minor', length=4, width=0.8) # Minor for months

            # 4. Vertical grid lines for major years only
            ax.grid(axis="x", alpha=0.3, which="major", linestyle="-", color="gray")

            # Mid-year dashed line (July) for visual guidance
            for d in self.dates:
                if d.month == 7:
                    ax.axvline(d, color="gray", alpha=0.1, linewidth=0.8, linestyle="--")

        else:
            # Broad view: Major every 5 years, minor every year
            ax.xaxis.set_major_locator(mdates.YearLocator(5))
            ax.xaxis.set_minor_locator(mdates.YearLocator(1))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))

            ax.tick_params(axis='x', which='major', length=7, width=1.2)
            ax.tick_params(axis='x', which='minor', length=0) 

            ax.grid(axis="x", alpha=0.2, which="major", linestyle="--")

        plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha="center")

    def plot_wealth_index(self, figsize=(16, 5), savepath=None):
        """
        Generates a Wealth Index chart color-coded by regime.
        """
        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)
    
        series = self.log_returns[:, 0].copy()
    
        # Auto-detection: price level vs log returns
        if np.abs(series).mean() > 1.0:
            # Price level -> normalize base 100
            wealth = 100 * series / series[0]
        else:
            # Log returns -> conversion to decimals and cumulative sum
            if np.abs(series).mean() > 0.1:
                series = series / 100.0
            wealth = 100 * np.exp(np.cumsum(series))
    
        ax.plot(self.dates, wealth, color="#2c3e50", linewidth=1.8,
                zorder=3, label=self.asset_names[0])
    
        ax.set_yscale("log")
        ax.yaxis.set_major_formatter(mticker.ScalarFormatter())
        ax.yaxis.set_minor_formatter(mticker.NullFormatter())
        ax.axhline(100, color="black", linewidth=0.5, linestyle="--", alpha=0.4)
        ax.set_ylabel("Wealth Index (Base 100, log scale)", fontsize=12)
        ax.set_title("Wealth Index Growth — by Regime", fontsize=14, fontweight="bold")
    
        # Combined legend
        patches = self._regime_patches()
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=patches + handles, loc="upper left", fontsize=12, framealpha=0.9, ncol=3)
    
        self._annotate_events(ax)
        self._format_xaxis(ax)
        ax.grid(axis="y", alpha=0.3, which="major")
        plt.tight_layout()
    
        if savepath:
            fig.savefig(savepath, dpi=150, bbox_inches="tight")
        plt.show()
        return fig

    def plot_implied_correlation(self, asset_i: int = 0, asset_j: int = 1, figsize: Tuple[int, int] = (16, 5),
                                 savepath: str = None):
        """
        Generates a chart of the implied correlation between two selected assets.

        Args:
            asset_i (int): Index of the first asset.
            asset_j (int): Index of the second asset.
        """
        if asset_i >= self.D or asset_j >= self.D:
            raise ValueError(f"Invalid indices. You have {self.D} assets (0 to {self.D-1}).")

        # 1. Retrieve probabilities
        probs = self.detector.regime_probabilities(series_index=0)

        # 2. Compute total conditional covariance at each step (T, D, D)
        total_covs = self.detector.conditional_covariance(probs, self.log_returns)

        # 3. Calculate implied correlation for the pair (i, j)
        vols = np.sqrt(np.diagonal(total_covs, axis1=-2, axis2=-1)) # (T, D)
        implied_corr = total_covs[:, asset_i, asset_j] / (vols[:, asset_i] * vols[:, asset_j] + 1e-16)

        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)

        label_pair = f"{self.asset_names[asset_i]} / {self.asset_names[asset_j]}"
        ax.plot(self.dates, implied_corr, color="#e74c3c", linewidth=2,
                label=f"{label_pair} Correlation", zorder=4)

        ax.axhline(0, color="black", linewidth=0.5)
        ax.set_ylabel("Correlation", fontsize=12)
        ax.set_title(f"Implied Correlation Dynamics: {label_pair}", fontsize=14, fontweight="bold")
        ax.set_ylim(-1.05, 1.05)

        # Combined legend
        patches = self._regime_patches()
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=patches + handles, loc="lower left", fontsize=12, framealpha=0.9, ncol=3)

        self._format_xaxis(ax)
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()

        if savepath: fig.savefig(savepath, dpi=150)
        plt.show()
        return fig
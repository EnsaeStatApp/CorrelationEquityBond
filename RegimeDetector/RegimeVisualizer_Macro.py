import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
import matplotlib.colors as mcolors
from matplotlib.patches import Patch
import numpy as np
import pandas as pd
from RegimeDetector.base import RegimeDetector
from typing import Union, List, Dict, Tuple


class RegimeVisualizer:
    """
    Classe dédiée à la visualisation temporelle des régimes.
    """
    def __init__(self, detector, Y, dates, states, asset_names=None,
                 regime_labels=None, regime_colors=None, events_dict=None, zoom_mode=False):
        self.detector = detector
        self.Y = Y
        self.dates = pd.to_datetime(dates)
        self.states = states
        self.zoom_mode = zoom_mode
        self.K = getattr(detector, 'K', len(np.unique(states)))
        self.D = Y.shape[1]
        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regime_labels = regime_labels or {k: f"Régime {k}" for k in range(self.K)}
        if regime_colors is not None:
            self.regime_colors = regime_colors
        else:
            cmap = plt.get_cmap('tab10')
            self.regime_colors = {k: mcolors.to_hex(cmap(k % 10)) for k in range(self.K)}
        self.events_dict = events_dict or {}

    def _shade_regimes(self, ax):
        for t in range(len(self.dates)):
            x0 = self.dates[t] - pd.Timedelta(days=15)
            x1 = self.dates[t] + pd.Timedelta(days=15)
            ax.axvspan(x0, x1, alpha=0.25,
                       color=self.regime_colors.get(self.states[t], "gray"),
                       linewidth=0, zorder=0)

    def _regime_patches(self):
        return [Patch(facecolor=self.regime_colors[k], alpha=0.5,
                      label=self.regime_labels.get(k, f"Régime {k}"))
                for k in sorted(self.regime_colors.keys()) if k < self.K]

    def _annotate_events(self, ax):
        if not self.events_dict:
            return
        for label, d in self.events_dict.items():
            d = pd.Timestamp(d)
            if self.dates.min() <= d <= self.dates.max():
                ax.axvline(d, color="#2c3e50", linewidth=0.8, alpha=0.4, linestyle=":")
                ymin, ymax = ax.get_ylim()
                if ax.get_yscale() == "log":
                    log_ypos = np.log10(ymin) + 0.92 * (np.log10(ymax) - np.log10(ymin))
                    ypos = 10 ** log_ypos
                else:
                    ypos = ymin + (ymax - ymin) * 0.92
                ax.annotate(label, xy=(d, ypos), fontsize=8, fontweight="bold",
                            ha="center", va="top", color="#2c3e50", alpha=0.8)

    def _format_xaxis(self, ax):
        if self.zoom_mode:
            ax.xaxis.set_major_locator(mdates.YearLocator())
            ax.xaxis.set_minor_locator(mdates.MonthLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
            ax.tick_params(axis='x', which='major', length=10, width=1.5, labelsize=10)
            ax.tick_params(axis='x', which='minor', length=4, width=0.8)
            ax.grid(axis="x", alpha=0.3, which="major", linestyle="-", color="gray")
            for d in self.dates:
                if d.month == 7:
                    ax.axvline(d, color="gray", alpha=0.1, linewidth=0.8, linestyle="--")
        else:
            ax.xaxis.set_major_locator(mdates.YearLocator(5))
            ax.xaxis.set_minor_locator(mdates.YearLocator(1))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
            ax.tick_params(axis='x', which='major', length=7, width=1.2)
            ax.tick_params(axis='x', which='minor', length=0)
            ax.grid(axis="x", alpha=0.2, which="major", linestyle="--")
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha="center")

    def plot_wealth_index(self, figsize=(16, 5), savepath=None):
        """
        Indice de richesse base 100 sur fond de régimes.
        Gère automatiquement niveaux de prix et log returns.
        """
        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)

        series = self.Y[:, 0].copy()

        # Niveau de prix → normalisation base 100
        if np.abs(series).mean() > 1.0:
            wealth = 100 * series / series[0]
        else:
            # Log returns en % → décimal
            if np.abs(series).mean() > 0.1:
                series = series / 100.0
            wealth = 100 * np.exp(np.cumsum(series))

        ax.plot(self.dates, wealth, color="#2c3e50", linewidth=1.8, zorder=3, label=self.asset_names[0])

        ax.set_yscale("log")
        ax.yaxis.set_major_formatter(mticker.ScalarFormatter())
        ax.yaxis.set_minor_formatter(mticker.NullFormatter())
        ax.axhline(100, color="black", linewidth=0.5, linestyle="--", alpha=0.4)

        ax.set_ylabel("Indice de richesse (base 100, log)", fontsize=12)
        ax.set_title("Indice de richesse — par régime", fontsize=14, fontweight="bold")

        patches = self._regime_patches()
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=patches + handles, loc="upper left", fontsize=9, framealpha=0.9, ncol=3)

        self._annotate_events(ax)
        self._format_xaxis(ax)
        ax.grid(axis="y", alpha=0.3, which="major")
        plt.tight_layout()

        if savepath:
            fig.savefig(savepath, dpi=150, bbox_inches="tight")
        plt.show()
        return fig

    def plot_implied_correlation(self, figsize=(16, 5), savepath=None):
        """
        Corrélation implicite SP500 / T-bond calculée depuis les données financières Y,
        conditionnellement aux régimes détectés sur les variables macro.
        """
        probs = self.detector.regime_probabilities()  # (T, K) — probabilités macro
        T, K = probs.shape
    
        # Covariances financières empiriques par régime (sur Y, pas sur les macro)
        states = self.states
        fin_covs = []
        for k in range(K):
            mask = (states == k)
            subset = self.Y[mask, :2]  # SP500 et T-bond uniquement
            if subset.shape[0] < 2:
                fin_covs.append(np.eye(2))
            else:
                fin_covs.append(np.cov(subset.T, ddof=1))
    
        # Corrélation implicite via loi des espérances totales
        implied_corr = np.zeros(T)
        for t in range(T):
            sigma_t = sum(probs[t, k] * fin_covs[k] for k in range(K))
            std_t = np.sqrt(np.diag(sigma_t))
            implied_corr[t] = sigma_t[0, 1] / (std_t[0] * std_t[1])
    
        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)
        ax.plot(self.dates, implied_corr, color="#e74c3c", linewidth=2,
                label="Corrélation Implicite du Modèle", zorder=4)
        ax.axhline(0, color="black", linewidth=0.5)
        ax.set_ylabel("Corrélation", fontsize=12)
        ax.set_title("Corrélation Stock-Bond via le HMM Macro", fontsize=14, fontweight="bold")
        ax.set_ylim(-1.05, 1.05)
    
        patches = self._regime_patches()
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=patches + handles, loc="lower left", fontsize=9, framealpha=0.9, ncol=3)
    
        self._format_xaxis(ax)
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()
    
        if savepath:
            fig.savefig(savepath, dpi=150, bbox_inches="tight")
        plt.show()
        return fig

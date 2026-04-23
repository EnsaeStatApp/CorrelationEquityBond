import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import Patch
import numpy as np
import pandas as pd
from typing import Union, List, Dict, Tuple
import matplotlib.dates as mdates


class RegimeReporter:
    """
    Classe dédiée à la visualisation temporelle des régimes
    """
    def __init__(self, detector, log_returns: np.ndarray, dates: Union[list, np.ndarray, pd.DatetimeIndex],
                 states: np.ndarray, asset_names: List[str] = None,
                 regime_labels: Dict[int, str] = None, regime_colors: List[str] = None,
                 events_dict: Dict[str, str] = None, zoom_mode: bool = False):
        """
        Paramètres :
        - detector : détecteur de régime (Market ou Macro) ayant déjà été fitté.
        - log_returns : tableau contenant les rendements logarithmiques des actifs financiers.
        - dates : index temporel correspondant aux lignes de log_returns et de states.
        - states : séquence des régimes détectés de taille (T,). Représente l'état dominant (Viterbi ou Filtré)
        - asset_names : noms des D actifs financiers (défaut : Actif i)
        - regime_labels : dictionnaire mappant l'identifiant du régime (int) à son nom descriptif (ex: {0: "Calme", 1: "Crise"}).
        - regime_colors : liste de codes couleurs pour chaque régime (doit être de taille au moins K).
        - events_dict : dictionnaire d'événements historiques à annoter sur les graphiques (ex: {"COVID-19": "2020-03-01"}).
        - zoom_mode : si true, zoom sur la période en question
        """
        self.detector = detector
        self.log_returns = log_returns
        self.dates = pd.to_datetime(dates)
        self.states = states
        self.zoom_mode = zoom_mode
        self.K = detector.K
        self.D = log_returns.shape[1]

        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regime_labels = regime_labels or {k: f"Régime {k}" for k in range(self.K)}

        # Couleurs harmonisées
        default_colors = ["#3498db", "#e74c3c", "#2ecc71", "#f1c40f", "#9b59b6"]
        self.regime_colors = {k: default_colors[k % len(default_colors)] for k in range(self.K)} if regime_colors is None else {k: regime_colors[k] for k in range(min(len(regime_colors), self.K))}

        self.events_dict = events_dict or {}

    def _shade_regimes(self, ax):
        """
        Ajoute le fond coloré selon la séquence d'états.
        Chaque span va du premier au dernier indice du régime,
        avec une extension minimale d'un mois pour les régimes d'un seul point.
        """
        n = len(self.dates)
        changes = np.where(self.states[:-1] != self.states[1:])[0]
        splits = np.concatenate(([0], changes + 1, [n]))
    
        for i in range(len(splits) - 1):
            start_idx = int(splits[i])
            end_idx = int(splits[i + 1] - 1)
            regime = self.states[start_idx]
    
            x0 = self.dates[start_idx]
    
            # Borne droite : date suivante si disponible, sinon +31 jours
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
        """"
        Génère les labels pour la légende dynamiquement selon K (méthode privée)
        """
        return [Patch(facecolor=color, alpha=0.5,
                      label=self.regime_labels.get(k, f"Régime {k}"))
                for k, color in self.regime_colors.items() if k < self.K]
    def _annotate_events(self, ax):
        """
        Ajoute les barres verticales pour les événements fournis (méthode privée)
        """
        if not self.events_dict:
            return

        for label, d in self.events_dict.items():
            d = pd.Timestamp(d)
            if self.dates.min() <= d <= self.dates.max():
                ax.axvline(d, color="#2c3e50", linewidth=0.8, alpha=0.4, linestyle=":")
                ymin, ymax = ax.get_ylim()
                # Gestion propre du positionnement sur un axe en log
                if ax.get_yscale() == "log":
                    log_ypos = np.log10(ymin) + 0.92 * (np.log10(ymax) - np.log10(ymin))
                    ypos = 10 ** log_ypos
                else:
                    ypos = ymin + (ymax - ymin) * 0.92
                ax.annotate(label, xy=(d, ypos), fontsize=8, fontweight="bold",
                            ha="center", va="top", color="#2c3e50", alpha=0.8)

    def _format_xaxis(self, ax):
        """
        Formatage fixe pour données mensuelles :
        - Majeur : Chaque année (Grand trait + Année)
        - Mineur : Chaque mois (Petit trait)
        - Spécifique : Milieu d'année (Juillet) via les paramètres de style
        """
        if self.zoom_mode:
            # 1. Positionnement des traits (Locators)
            # Majeur = 1er Janvier
            ax.xaxis.set_major_locator(mdates.YearLocator())
            # Mineur = Chaque mois
            ax.xaxis.set_minor_locator(mdates.MonthLocator())

            # 2. Format du texte (Formatter)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))

            # 3. Style des traits (Ticks)
            ax.tick_params(axis='x', which='major', length=10, width=1.5, labelsize=10) # Grands traits pour les années
            ax.tick_params(axis='x', which='minor', length=4, width=0.8) # Petits traits pour les mois

            # 4. On ajoute les traits de grille verticaux pour les années uniquement
            ax.grid(axis="x", alpha=0.3, which="major", linestyle="-", color="gray")

            # Pour le "trait moyen" au milieu de l'année (Juillet) :
            ax.xaxis.set_minor_locator(mdates.MonthLocator()) # On garde tous les mois en ticks

            # Trait vertical au milieu :
            for d in self.dates:
                if d.month == 7:
                    ax.axvline(d, color="gray", alpha=0.1, linewidth=0.8, linestyle="--")

        else:
            ax.xaxis.set_major_locator(mdates.YearLocator(5))
            ax.xaxis.set_minor_locator(mdates.YearLocator(1))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))

            # Traits plus sobres
            ax.tick_params(axis='x', which='major', length=7, width=1.2)
            ax.tick_params(axis='x', which='minor', length=0) # Pas de traits pour les mois

            # Grille tous les 5 ans
            ax.grid(axis="x", alpha=0.2, which="major", linestyle="--")

        plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha="center")



    def plot_wealth_index(self, figsize=(16, 5), savepath=None):
        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)
    
        series = self.log_returns[:, 0].copy()
    
        # Détection automatique : niveau de prix vs log returns
        if np.abs(series).mean() > 1.0:
            # Niveau de prix → normalisation base 100
            wealth = 100 * series / series[0]
        else:
            # Log returns en % → conversion décimal
            if np.abs(series).mean() > 0.1:
                series = series / 100.0
            wealth = 100 * np.exp(np.cumsum(series))
    
        ax.plot(self.dates, wealth, color="#2c3e50", linewidth=1.8,
                zorder=3, label=self.asset_names[0])
    
        ax.set_yscale("log")
        ax.yaxis.set_major_formatter(mticker.ScalarFormatter())
        ax.yaxis.set_minor_formatter(mticker.NullFormatter())
        ax.axhline(100, color="black", linewidth=0.5, linestyle="--", alpha=0.4)
        ax.set_ylabel("Indice de richesse (base 100, log)", fontsize=12)
        ax.set_title("Indice de richesse — par régime", fontsize=14, fontweight="bold")
    
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
        Génère le graphique de la corrélation entre deux actifs choisis
        """
        if asset_i >= self.D or asset_j >= self.D:
            raise ValueError(f"Indices invalides. Vous avez {self.D} actifs (0 à {self.D-1}).")

        # 1. On récupère les probas
        probs = self.detector.regime_probabilities(series_index=0)

        # 2. On calcule la covariance totale à chaque instant t (T, D, D)
        total_covs = self.detector.conditional_covariance(probs, self.log_returns)

        # 3. Calcul de la corrélation impliquée pour la paire (i, j)
        vols = np.sqrt(np.diagonal(total_covs, axis1=-2, axis2=-1)) # (T, D)
        implied_corr = total_covs[:, asset_i, asset_j] / (vols[:, asset_i] * vols[:, asset_j] + 1e-16)

        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)

        label_pair = f"{self.asset_names[asset_i]} / {self.asset_names[asset_j]}"
        ax.plot(self.dates, implied_corr, color="#e74c3c", linewidth=2,
                label=f"Corrélation {label_pair}", zorder=4)

        ax.axhline(0, color="black", linewidth=0.5)
        ax.set_ylabel("Corrélation", fontsize=12)
        ax.set_title(f"Dynamique de Corrélation Impliquée : {label_pair}", fontsize=14, fontweight="bold")
        ax.set_ylim(-1.05, 1.05)

        # Légende combinée
        patches = self._regime_patches()
        handles, _ = ax.get_legend_handles_labels()
        ax.legend(handles=patches + handles, loc="lower left", fontsize=12, framealpha=0.9, ncol=3)

        self._format_xaxis(ax)
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()

        if savepath: fig.savefig(savepath, dpi=150)
        plt.show()
        return fig

    

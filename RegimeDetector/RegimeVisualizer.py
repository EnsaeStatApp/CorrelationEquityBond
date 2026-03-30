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
    Classe dédiée à la visualisation temporelle des régimes
    """
    def __init__(self, detector : RegimeDetector, Y : np.ndarray, dates : Union[list, np.ndarray], 
                 states : np.ndarray, asset_names : List[str] = None, regime_labels : Dict[int, str] = None, 
                 regime_colors: Dict[int, str] = None, events_dict :Dict[str, str] = None, zoom_mode : bool = False):
        """
        Paramètres : 
        - detector : un détecteur de régimes 
        - Y : tableaux de log returns
        - dates : les dates de Y
        - states : tableaux des régimes les + probables
        - asset_names : noms des actifs (défaut : "Asset i )
        - regime_labels : nom des régimes (à déterminer après les avoir analysé, défaut : "Régime i")
        - regime_colors : couleurs associées à chaque régime (sinon par défaut)
        - events_dict : dictionnaire d'évenements macroéconomiques à afficher sur le graphique
        - zoom_mode : True si on veut zoomer sur une période particulière (et regarder mois par mois), sinon False
        """
        
        self.detector = detector  
        self.Y = Y
        self.dates = pd.to_datetime(dates)
        self.states = states
        self.zoom_mode = zoom_mode
        
        # Récupération dynamique de K
        self.K = getattr(detector, 'K', len(np.unique(states)))
        self.D = Y.shape[1]
        
        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regime_labels = regime_labels or {k: f"Régime {k}" for k in range(self.K)}
        
        # Gestion dynamique des couleurs (palette auto si non fournie)
        if regime_colors is not None:
            self.regime_colors = regime_colors
        else:
            cmap = plt.get_cmap('tab10')
            self.regime_colors = {k: mcolors.to_hex(cmap(k % 10)) for k in range(self.K)}
            
        # Dictionnaire d'événements optionnel (ex: {"Crise": "2008-09-01"})
        self.events_dict = events_dict or {}

    def _shade_regimes(self, ax):
        """
        Ajoute le fond coloré selon la séquence d'états (méthode privée)
        """
        for t in range(len(self.dates)):
            x0 = self.dates[t] - pd.Timedelta(days=15)
            x1 = self.dates[t] + pd.Timedelta(days=15)
            ax.axvspan(x0, x1, alpha=0.25,
                       color=self.regime_colors.get(self.states[t], "gray"),
                       linewidth=0, zorder=0)

    def _regime_patches(self):
        """
        Génère les labels pour la légende dynamiquement selon K (méthode privée)
        """
        return [Patch(facecolor=self.regime_colors[k], alpha=0.5,
                      label=self.regime_labels.get(k, f"Régime {k}"))
                for k in sorted(self.regime_colors.keys()) if k < self.K]

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



    def plot_wealth_index(self, figsize : Tuple[int, int]=(16, 5), savepath : str = None):
        """
        Génère le graphique de la performance cumulée du S&P sur fond de régimes

        Paramètres : 
        - figsize : taille de la figure
        - savepath : chemin pour enregistrer l'image
        """
        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)
        log_ret = self.Y[:, 0]
        line_colors = ["#2c3e50", "#8e44ad", "#16a085", "#d35400"]
        
        # on ne trace que pour S&P. TO DO : vérifier que c'est bien l'index 0
        scale_factor = 1.0
        if np.abs(log_ret).mean() > 0.1: # echelle pourcentage donc on va diviser pour l'échelle
            scale_factor = 100.0
        wealth = 100 * np.exp(np.cumsum(self.Y[:, 0])/scale_factor) 
        ax.plot(self.dates, wealth, color=line_colors[0 % len(line_colors)],
                linewidth=1.8, zorder=3, label=self.asset_names[0])

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
        Génère le graphique de la corrélation entre les 2 premiers actifs (S&P et T-bond 10y) 

        Paramètres : 
        - figsize : taille de la figure
        - savepath : chemin pour enregistrer l'image
        """
        probs = self.detector.regime_probabilities() # (T, K)
        
        T = len(probs)
        implied_corr = np.zeros(T)
        
        for t in range(T):
            # Covariance totale au temps t
            sigma_t = self.detector.conditional_covariance(t)
            
            # Corrélation extraite de sigma_t
            std_t = np.sqrt(np.diag(sigma_t))
            implied_corr[t] = sigma_t[0, 1] / (std_t[0] * std_t[1])
    
        fig, ax = plt.subplots(figsize=figsize)
        self._shade_regimes(ax)

        ax.plot(self.dates, implied_corr, color="#e74c3c", linewidth=2, label="Corrélation Implicite du Modèle", zorder=4)

        ax.axhline(0, color="black", linewidth=0.5)
        ax.set_ylabel("Corrélation", fontsize=12)
        ax.set_title("Corrélation Stock-Bond via le HMM", fontsize=14, fontweight="bold")
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

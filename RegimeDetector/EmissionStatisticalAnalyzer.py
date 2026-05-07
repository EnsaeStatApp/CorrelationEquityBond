from statsmodels.stats.multitest import multipletests
import numpy as np
import pandas as pd
from scipy import stats
from itertools import combinations
from typing import List


class EmissionStatisticalAnalyzer:
    """
    Classe générique qui effectue des tests statistiques descriptifs sur des tableaux de log_returns selon leur régime

    - Renvoie les statistiques descriptives propres à chaque régime
    - Applique des tests de Levene entre deux régimes pour déterminer si leur volatilité sur un actif diffère significativement
    - Applique des tests de Fisher-Transfo entre deux régimes pour déterminer si la corrélation entre deux actifs diffère significativement
    """
    def __init__(self, log_returns: np.ndarray, states: np.ndarray, asset_names: List = None):
        """
        Parametres:
        - Y : tableaux de log returns sur laquelle on va faire les tests stats
        - states : tableaux de régimes sur chaque date
        - asset_names : nom des actifs
        """
        self.log_returns = log_returns
        self.states = states
        self.D = log_returns.shape[1]
        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regimes = np.unique(states)

    @staticmethod
    def _safe_corrcoef(Y: np.ndarray) -> np.ndarray:
        """
        Calcule la matrice de corrélation en gérant le cas où un actif a une std nulle
        (régime avec trop peu d'observations identiques). Les paires impliquant un actif
        à variance nulle reçoivent NaN au lieu de provoquer un RuntimeWarning.
        """
        with np.errstate(invalid='ignore', divide='ignore'):
            C = np.corrcoef(Y, rowvar=False)
        # Remplacer les NaN diagonaux par 1 (corrélation d'un actif avec lui-même)
        np.fill_diagonal(C, 1.0)
        return C

    def descriptive_stats(self):
        """
        Renvoie la description purement statistiques de chacun des régimes (i.e. un régime est décrit par les stats de chacun des actifs dans celui-ci)
        """
        summary = {}
        for k in self.regimes:
            mask = self.states == k
            Yk = self.log_returns [mask] # rendements du régime k
            n = Yk.shape[0]
            summary[k] = { # stats descriptives
                "n": n,
                "mean": np.mean(Yk, axis=0),
                "std": np.std(Yk, axis=0, ddof=1),
                "corr": self._safe_corrcoef(Yk),
                "cov": np.cov(Yk, rowvar=False),
                "asset_names": self.asset_names,
            }
        return summary

    def get_descriptive_stats_df(self, annualise=np.sqrt(12)):
        """
        Renvoie la description purement statistiques sous forme de df pour un affichage propre à la fin
        """
        summ = self.descriptive_stats()
        rows = []
        for k, s in summ.items():
            for d in range(self.D):
                rows.append({
                    "Regime": k,
                    "Asset": self.asset_names[d],
                    "N": s["n"],
                    "Mean": s["mean"][d],
                    "Vol": s["std"][d],
                    "Ann_Vol": s["std"][d] * annualise
                })
        return pd.DataFrame(rows)

    def _apply_correction(self, df, correction : str = "bonferroni", alpha : float = 0.05):
        """
        Applique une correction de manière à mieux refleter les significativités

        Paramètres:
        - correction : type de la correction souhaité (bonferroni ou fdr)
        - alpha : seuil de significativité
        """
        if len(df) == 0:
            return df

        pvals = df["p_value"].values

        # Les lignes avec NaN (régimes trop petits pour le test) sont exclues de la correction
        # puis réintégrées avec p_adjusted=NaN et significant=False
        valid_mask = ~np.isnan(pvals)
        pvals_corrected = np.full(len(pvals), np.nan)
        reject = np.zeros(len(pvals), dtype=bool)

        if valid_mask.sum() > 0:
            pvals_valid = pvals[valid_mask]
            if correction == "bonferroni":
                rej_v, pvals_v_corr, _, _ = multipletests(pvals_valid, alpha=alpha, method='bonferroni')
            elif correction == "fdr":
                rej_v, pvals_v_corr, _, _ = multipletests(pvals_valid, alpha=alpha, method='fdr_bh')
            else:
                pvals_v_corr = pvals_valid
                rej_v = pvals_valid < alpha

            pvals_corrected[valid_mask] = pvals_v_corr
            reject[valid_mask] = rej_v

        df["p_adjusted"] = pvals_corrected
        df["significant"] = reject

        return df

    def test_pairwise_levene(self, correction : str = "bonferroni", alpha : float = 0.05):
        """
        Vérifie qu'entre chaque régime i et j les volatilités sont bien distinctes pour un actif a_k
        H0 : Var(a_k | régime = i) = Var(a_k | régime = j)

        Paramètres:
        - correction : type de la correction souhaité (bonferroni ou fdr)
        - alpha : seuil de significativité
        """
        rows = []
        for (k1, k2) in combinations(self.regimes, 2):
            for d in range(self.D):
                g1 = self.log_returns [self.states == k1, d] # rendement de l'actif D dans le régime k1
                g2 = self.log_returns [self.states == k2, d] # # rendement de l'actif D dans le régime k2
                stat, pval = stats.levene(g1, g2, center="median") # on centre sur la médiane à cause des valeurs extrèmes
                rows.append({
                "regime_pair": f"{k1} vs {k2}",
                "asset": self.asset_names[d],
                "vol_A": round(np.std(g1, ddof=1), 4), # Nom fixe
                "vol_B": round(np.std(g2, ddof=1), 4), # Nom fixe
                "levene_stat": round(stat, 4),
                "p_value": pval,
            })
        return self._apply_correction(pd.DataFrame(rows), correction, alpha) # on applique la correction des p-values

    def get_summary_table(self, n_boot : int = 5000, alpha : float = 0.05, annualise : float = np.sqrt(12)):
        """
        Regroupe Moyenne, Vol, et IC Bootstrap dans un seul tableau lisible

        Paramètres :
        - n_boot : nombres d'itérations du bootstrap
        - alpha : seuil de significativité
        - annualise : facteur d'annualisation
        """
        summ = self.descriptive_stats()
        boot = self.generate_bootstrap_ci(n_boot=n_boot, alpha=alpha)
        rows = []

        for k in self.regimes:
            for d, asset in enumerate(self.asset_names):
                ci = boot[k]["vol_ci"][asset]

                rows.append({
                    "Regime": k,
                    "Asset": asset,
                    "Mean": summ[k]["mean"][d],
                    "Vol_Ann": summ[k]["std"][d] * annualise,
                    "CI_Low": ci["ci_low"] * annualise,
                    "CI_High": ci["ci_high"] * annualise
                })

        return pd.DataFrame(rows).round(4).set_index(["Regime", "Asset"]) # index par régime pour lisibilité


    def test_fisher_z(self, correction="bonferroni", alpha : float = 0.05):
        """
        Vérifie qu'entre chaque régime i et j les corrélations sont bien distinctes entre deux actifs a_k et a_l
        H0 : Corr(a_k, a_l | régime = i) = Corr(a_k, a_l | régime = j)

        Paramètres:
        - correction : type de la correction souhaité (bonferroni ou fdr)
        - alpha : seuil de significativité
        """

        def _fisher_z(r : float):
            """
            Calcule la transformer de Fisher pour rendre la corrélation "plus normale"

            Paramètre :
            - r = corrélation
            """
            r = np.clip(r, -0.9999, 0.9999) # on clip au cas où
            return np.arctanh(r)

        def fisher_z_test(r1 : float, n1 : int, r2 : float, n2 : int):
            """
            Compare les 2 corrélations issues des deux régimes différents

            Paramètres :
            - r1 : corrélation du régime 1
            - n1 : taille du régime 1
            - r2 : corrélation du régime 2
            - n2 : taille du régime 2
            """
            # Le test requiert n > 3 pour chaque régime (dénominateur 1/(n-3))
            # Si un régime est trop peu peuplé, le test n'est pas calculable → NaN
            if n1 <= 3 or n2 <= 3:
                return np.nan, np.nan
            z1 = _fisher_z(r1)
            z2 = _fisher_z(r2)
            se = np.sqrt(1.0 / (n1 - 3) + 1.0 / (n2 - 3)) # erreur standard
            z_stat = (z1 - z2) / se # statistique Z
            p_value = 2 * stats.norm.sf(np.abs(z_stat)) # Z suit une loi normale
            return z_stat, p_value

        summ = self.descriptive_stats() # utilise la fonction de description des stats par régimes
        rows = []

        for (k1, k2) in combinations(self.regimes, 2):
            n1, n2 = summ[k1]["n"], summ[k2]["n"]
            for i, j in combinations(range(self.D), 2):
                r1 = summ[k1]["corr"][i, j]
                r2 = summ[k2]["corr"][i, j]
                z_stat, pval = fisher_z_test(r1, n1, r2, n2)
                rows.append({
                    "regime_pair": f"{k1} vs {k2}",
                    "asset_pair": f"{self.asset_names[i]} / {self.asset_names[j]}",
                    "corr_A": r1, # Nom générique
                    "corr_B": r2, # Nom générique
                    "fisher_z": z_stat,
                    "p_value": pval,
                })
        return self._apply_correction(pd.DataFrame(rows), correction, alpha)

    def generate_bootstrap_ci(self, n_boot : int = 5000, alpha : float = 0.05, seed : int = 42):
        """
        Applique un bootstrap non-paramétrique pour avoir des intervalles de confiance des statistiques de régimes

        Paramètres :
        - n_boot : nombres d'intérations du bootstrap (défault : 5000)
        - alpha : niveau de confiance
        - seed : seed surlequel est fait le bootstrap
        """
        rng = np.random.RandomState(seed)
        boot_results = {}

        for k in self.regimes:
            Yk = self.log_returns [self.states == k] # on créé des sous-échantillons qui représentent les log returns dans chacun des régimes
            nk = Yk.shape[0] # n_k = nombres d'observations dans le régime k

            vol_samples = np.empty((n_boot, self.D)) # va stocker les volatilités des n_boot itérations
            pair_keys = [(i, j) for i, j in combinations(range(self.D), 2)]
            corr_samples = np.empty((n_boot, len(pair_keys))) # va stocker les corrélations des n_boot itérations

            for b in range(n_boot):
                idx = rng.choice(nk, size=nk, replace=True) # on pioche n_k observations dans le régime k avec remise
                Yb = Yk[idx] # les logs returns pour chaque tirage
                vol_samples[b] = np.std(Yb, axis=0, ddof=1) # vol divisée par n-1
                C = self._safe_corrcoef(Yb) # matrice de correl
                for p_idx, (i, j) in enumerate(pair_keys):
                    corr_samples[b, p_idx] = C[i, j]

            lo = alpha / 2 * 100 # borne basse
            hi = (1 - alpha / 2) * 100 # borne haute

            # volatilités
            vol_ci = {}
            for d in range(self.D):
                point = np.std(Yk, axis=0, ddof=1)[d]
                vol_ci[self.asset_names[d]] = {
                    "point": point,
                    "ci_low": np.percentile(vol_samples[:, d], lo),
                    "ci_high": np.percentile(vol_samples[:, d], hi),
                }

            # corréls
            corr_ci = {}
            C_full = self._safe_corrcoef(Yk)
            for p_idx, (i, j) in enumerate(pair_keys):
                label = f"{self.asset_names[i]} / {self.asset_names[j]}"
                corr_ci[label] = {
                    "point": C_full[i, j],
                    "ci_low": np.percentile(corr_samples[:, p_idx], lo),
                    "ci_high": np.percentile(corr_samples[:, p_idx], hi),
                }

            boot_results[k] = {"vol_ci": vol_ci, "corr_ci": corr_ci}

        return boot_results


    def get_bootstrap_ci_df(self, n_boot : int = 5000, alpha : float = 0.05, seed : int = 42, annualise : float = np.sqrt(12)):
        """
        Renvoie les résultats du bootstrap sous forme de df pour un affichage propre à la fin
        """
        boot_data = self.generate_bootstrap_ci(n_boot=n_boot, alpha=alpha, seed=seed)
        rows = []

        # on met un "label" commun pour regrouper ensemble
        for k, metrics in boot_data.items():
            # 1. volatilités
            for asset, ci in metrics["vol_ci"].items():
                rows.append({
                    "Regime": k,
                    "Type": "Volatility",
                    "Label": asset,
                    "Value": ci["point"] * annualise,
                    "Lower_Bound": ci["ci_low"] * annualise,
                    "Upper_Bound": ci["ci_high"] * annualise
                })

            # 2. corrélations
            for pair, ci in metrics["corr_ci"].items():
                rows.append({
                    "Regime": k,
                    "Type": "Correlation",
                    "Label": pair,
                    "Value": ci["point"],
                    "Lower_Bound": ci["ci_low"],
                    "Upper_Bound": ci["ci_high"]
                })

        return pd.DataFrame(rows).set_index(["Regime", "Type", "Label"])


    

from statsmodels.stats.multitest import multipletests
import numpy as np
import pandas as pd
from scipy import stats
from itertools import combinations


class RegimeStatisticalAnalyzer:
    """
    Classe générique qui effectue des tests statistiques descriptifs sur des tableaux selon leur régime
    """
    def __init__(self, Y: np.ndarray, states: np.ndarray, asset_names: list = None):
        """
        Parametres: 
        - Y : tableaux de log returns sur laquelle on va faire les tests stats
        - states : tableaux de régimes sur chaque date
        - asset_names : nom des actifs
        """
        self.Y = Y
        self.states = states
        self.D = Y.shape[1]
        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regimes = np.unique(states)

    def descriptive_stats(self):
        """
        Renvoie la description purement statistiques de chacun des régimes (i.e. un régime est décrit par les stats de chacun des actifs dans celui-ci)
        """
        summary = {}
        for k in self.regimes:
            mask = self.states == k
            Yk = self.Y[mask] # rendements du régime k 
            n = Yk.shape[0]
            summary[k] = { # stats descriptives
                "n": n,
                "mean": np.mean(Yk, axis=0),
                "std": np.std(Yk, axis=0, ddof=1),
                "corr": np.corrcoef(Yk, rowvar=False),
                "cov": np.cov(Yk, rowvar=False),
                "asset_names": self.asset_names,
            }
        return summary
    
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

        if correction == "bonferroni": # pour bonferroni, il faut que la p-value soit inf à alpha/nbr de tests
            reject, pvals_corrected, _, _ = multipletests(pvals, alpha=alpha, method='bonferroni')
        elif correction == "fdr":  # pour fdr, la p-value corrigée devient p_old * nbr de tests/rang
            reject, pvals_corrected, _, _ = multipletests(pvals, alpha=alpha, method='fdr_bh')
        else:
            pvals_corrected = pvals
            reject = pvals < alpha

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
                g1 = self.Y[self.states == k1, d] # rendement de l'actif D dans le régime k1
                g2 = self.Y[self.states == k2, d] # # rendement de l'actif D dans le régime k2
                stat, pval = stats.levene(g1, g2, center="median") # on centre sur la médiane à cause des valeurs extrèmes
                rows.append({
                    "regime_pair": f"{k1} vs {k2}",
                    "asset": self.asset_names[d],
                    f"vol_regime_{k1}": round(np.std(g1, ddof=1), 4), # ddof = 1 pour diviser par n-1
                    f"vol_regime_{k2}": round(np.std(g2, ddof=1), 4),
                    "levene_stat": round(stat, 4),
                    "p_value": pval,
                })
        return self._apply_correction(pd.DataFrame(rows), correction, alpha) # on applique la correction des p-values

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
                    "corr_regime_{:d}".format(k1): r1,
                    "corr_regime_{:d}".format(k2): r2,
                    "fisher_z": z_stat,
                    "p_value": pval,
                })
        return self._apply_correction(pd.DataFrame(rows), correction, alpha)

    def generate_bootstrap_ci(self, n_boot : int = 5000, alpha : float = 0.05, seed : int = 42):
        """
        Applique un bootstrap paramétrique pour avoir des intervalles de confiance des statistiques de régimes

        Paramètres : 
        - n_boot : nombres d'intérations du bootstrap (défault : 5000)
        - alpha : niveau de confiance
        - seed : seed surlequel est fait le bootstrap
        """
        rng = np.random.RandomState(seed)
        boot_results = {}

        for k in self.regimes:
            Yk = self.Y[self.states == k] # rendements du régime k 
            nk = Yk.shape[0]

            vol_samples = np.empty((n_boot, self.D)) # va stocker les volatilités des n_boot itérations
            pair_keys = [(i, j) for i, j in combinations(range(self.D), 2)]
            corr_samples = np.empty((n_boot, len(pair_keys))) # va stocker les corrélations des n_boot itérations

            for b in range(n_boot):
                idx = rng.choice(nk, size=nk, replace=True) # tire au sort un indice de régime puis le remet dans les tirages
                Yb = Yk[idx]
                vol_samples[b] = np.std(Yb, axis=0, ddof=1) # vol divisée par n-1
                C = np.corrcoef(Yb, rowvar=False) # matrice de correl
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
            C_full = np.corrcoef(Yk, rowvar=False)
            for p_idx, (i, j) in enumerate(pair_keys):
                label = f"{self.asset_names[i]} / {self.asset_names[j]}"
                corr_ci[label] = {
                    "point": C_full[i, j],
                    "ci_low": np.percentile(corr_samples[:, p_idx], lo),
                    "ci_high": np.percentile(corr_samples[:, p_idx], hi),
                }

            boot_results[k] = {"vol_ci": vol_ci, "corr_ci": corr_ci}

        return boot_results

    def full_report(self, n_boot : int = 5000, alpha : float = 0.05, annualise : float = np.sqrt(12), correction : str = "bonferroni", seed : int = 42):
        """
        Propose un affichage propre et lisible des résultats de tous les tests statistiques

        Paramètres : 
        - n_boot : nombres d'intérations du bootstrap (défault : 5000)
        - alpha : niveau de confiance
        - annualise : facteur de multiplication pour avoir des métriques annualisées
        - seed : seed surlequel est fait le bootstrap

        """
        summ = self.descriptive_stats()
        df_levene = self.test_pairwise_levene(correction=correction)
        df_fisher = self.test_fisher_z(correction=correction)
        boot = self.generate_bootstrap_ci(n_boot=n_boot, alpha=alpha, seed=seed)

        print("=" * 70)
        print("STATISTIQUES DESCRIPTIVES PAR RÉGIME")
        print("=" * 70)
        for k in self.regimes:
            s = summ[k]
            print(f"\n--- Régime {k}  (n = {s['n']}) ---")
            for d in range(self.D):
                ann_vol = s["std"][d] * annualise
                print(f"  {self.asset_names[d]:>25s}:  moy = {s['mean'][d]:+.4f}   "
                      f"vol = {s['std'][d]:.4f}  (ann. {ann_vol:.4f})")
            print(f"  Matrice de corrélation:\n{np.array2string(s['corr'], precision=4, suppress_small=True)}")

        print("\n" + "=" * 70)
        print(f"TEST DE LEVENE (VARIANCES) — correction: {correction}")
        print("=" * 70)
        if not df_levene.empty:
            print(df_levene.to_string(index=False))
        else:
            print("  (Besoin d'au moins 2 régimes)")

        print("\n" + "=" * 70)
        print(f"TEST DE FISHER Z (CORRÉLATIONS) — correction: {correction}")
        print("=" * 70)
        if not df_fisher.empty:
            print(df_fisher.to_string(index=False))
        else:
            print("  (Besoin d'au moins 2 régimes et 2 actifs)")

        print("\n" + "=" * 70)
        print(f"INTERVALLES DE CONFIANCE BOOTSTRAP ({(1-alpha)*100:.0f}%)  ({n_boot} tirages)")
        print("=" * 70)
        for k in self.regimes:
            print(f"\n--- Régime {k} ---")
            for name, v in boot[k]["vol_ci"].items():
                print(f"  Vol  {name:>25s}: {v['point']*annualise:.4f}  "
                      f"[{v['ci_low']*annualise:.4f}, {v['ci_high']*annualise:.4f}]")
            for name, c in boot[k]["corr_ci"].items():
                print(f"  Corr {name:>25s}: {c['point']:+.4f}  "
                      f"[{c['ci_low']:+.4f}, {c['ci_high']:+.4f}]")

        return {
            "descriptive": summ,
            "levene": df_levene,
            "fisher": df_fisher,
            "bootstrap": boot,
        }

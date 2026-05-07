from statsmodels.stats.multitest import multipletests
import numpy as np
import pandas as pd
from scipy import stats
from itertools import combinations
from typing import List


class EmissionStatisticalAnalyzer:
    """
    Generic class that performs descriptive statistical tests on log_returns based on their regime.

    - Returns descriptive statistics specific to each regime.
    - Applies Levene's tests between two regimes to determine if asset volatility differs significantly.
    - Applies Fisher-Transformation tests between two regimes to determine if asset correlations differ significantly.
    """

    def __init__(self, log_returns: np.ndarray, states: np.ndarray, asset_names: List = None):
        """
        Parameters:
        -----------
        log_returns : np.ndarray
            Array of log returns to be analyzed.
        states : np.ndarray
            Array of regimes/states corresponding to each date.
        asset_names : List, optional
            Names of the assets. Defaults to generic naming if None.
        """
        self.log_returns = log_returns
        self.states = states
        self.D = log_returns.shape[1]
        self.asset_names = asset_names or [f"Asset_{i}" for i in range(self.D)]
        self.regimes = np.unique(states)

    @staticmethod
    def _safe_corrcoef(Y: np.ndarray) -> np.ndarray:
        """
        Calculates the correlation matrix while handling cases where an asset has zero standard deviation 
        (e.g., regimes with too few or identical observations). Pairs involving a zero-variance asset 
        receive NaN instead of triggering a RuntimeWarning.
        """
        with np.errstate(invalid='ignore', divide='ignore'):
            C = np.corrcoef(Y, rowvar=False)
        # Replace diagonal NaNs with 1.0 (correlation of an asset with itself)
        np.fill_diagonal(C, 1.0)
        return C

    def descriptive_stats(self):
        """
        Returns purely statistical descriptions of each regime (i.e., a regime is described 
        by the statistics of each asset within it).
        """
        summary = {}
        for k in self.regimes:
            mask = self.states == k
            Yk = self.log_returns[mask]  # Returns for regime k
            n = Yk.shape[0]
            summary[k] = {  # Descriptive stats
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
        Returns descriptive statistics as a DataFrame for clean display.
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

    def _apply_correction(self, df, correction: str = "bonferroni", alpha: float = 0.05):
        """
        Applies p-value correction methods to better reflect statistical significance.

        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame containing a 'p_value' column.
        correction : str
            Type of correction requested ("bonferroni" or "fdr").
        alpha : float
            Significance threshold.
        """
        if len(df) == 0:
            return df

        pvals = df["p_value"].values

        # Rows with NaN (regimes too small for the test) are excluded from correction,
        # then reintegrated with p_adjusted=NaN and significant=False.
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

    def test_pairwise_levene(self, correction: str = "bonferroni", alpha: float = 0.05):
        """
        Checks if volatilities are significantly distinct between regimes i and j for each asset a_k.
        H0: Var(a_k | regime = i) = Var(a_k | regime = j)

        Parameters:
        -----------
        correction : str
            Type of p-value correction ("bonferroni" or "fdr").
        alpha : float
            Significance threshold.
        """
        rows = []
        for (k1, k2) in combinations(self.regimes, 2):
            for d in range(self.D):
                g1 = self.log_returns[self.states == k1, d]  # Asset D returns in regime k1
                g2 = self.log_returns[self.states == k2, d]  # Asset D returns in regime k2
                # Centered on median to handle potential outliers
                stat, pval = stats.levene(g1, g2, center="median")
                rows.append({
                    "regime_pair": f"{k1} vs {k2}",
                    "asset": self.asset_names[d],
                    "vol_A": round(np.std(g1, ddof=1), 4),
                    "vol_B": round(np.std(g2, ddof=1), 4),
                    "levene_stat": round(stat, 4),
                    "p_value": pval,
                })
        # Apply p-value correction
        return self._apply_correction(pd.DataFrame(rows), correction, alpha)

    def get_summary_table(self, n_boot: int = 5000, alpha: float = 0.05, annualise: float = np.sqrt(12)):
        """
        Consolidates Mean, Volatility, and Bootstrap Confidence Intervals into a single readable table.

        Parameters:
        -----------
        n_boot : int
            Number of bootstrap iterations.
        alpha : float
            Significance threshold.
        annualise : float
            Annualization factor.
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

        # Set index by regime for readability
        return pd.DataFrame(rows).round(4).set_index(["Regime", "Asset"])

    def test_fisher_z(self, correction="bonferroni", alpha: float = 0.05):
        """
        Checks if correlations between two assets (a_k, a_l) are significantly distinct 
        between regimes i and j.
        H0: Corr(a_k, a_l | regime = i) = Corr(a_k, a_l | regime = j)

        Parameters:
        -----------
        correction : str
            Type of p-value correction ("bonferroni" or "fdr").
        alpha : float
            Significance threshold.
        """

        def _fisher_z(r: float):
            """
            Calculates the Fisher Z-transformation to normalize the correlation distribution.

            Parameter:
            ----------
            r : float
                Correlation coefficient.
            """
            r = np.clip(r, -0.9999, 0.9999)  # Safety clip
            return np.arctanh(r)

        def fisher_z_test(r1: float, n1: int, r2: float, n2: int):
            """
            Compares two correlations from different regimes.

            Parameters:
            -----------
            r1, r2 : float
                Correlations for regimes 1 and 2.
            n1, n2 : int
                Sample sizes for regimes 1 and 2.
            """
            # Test requires n > 3 for each regime (due to the 1/(n-3) denominator).
            # If a regime is underpopulated, the test returns NaN.
            if n1 <= 3 or n2 <= 3:
                return np.nan, np.nan
            z1 = _fisher_z(r1)
            z2 = _fisher_z(r2)
            se = np.sqrt(1.0 / (n1 - 3) + 1.0 / (n2 - 3))  # Standard error
            z_stat = (z1 - z2) / se  # Z-statistic
            p_value = 2 * stats.norm.sf(np.abs(z_stat))  # Z follows a normal distribution
            return z_stat, p_value

        summ = self.descriptive_stats()
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
                    "corr_A": r1,
                    "corr_B": r2,
                    "fisher_z": z_stat,
                    "p_value": pval,
                })
        return self._apply_correction(pd.DataFrame(rows), correction, alpha)

    def generate_bootstrap_ci(self, n_boot: int = 5000, alpha: float = 0.05, seed: int = 42):
        """
        Applies non-parametric bootstrap to obtain confidence intervals for regime statistics.

        Parameters:
        -----------
        n_boot : int
            Number of bootstrap iterations (default: 5000).
        alpha : float
            Confidence level.
        seed : int
            Random seed for reproducibility.
        """
        rng = np.random.RandomState(seed)
        boot_results = {}

        for k in self.regimes:
            # Create sub-samples representing log returns in each regime
            Yk = self.log_returns[self.states == k]
            nk = Yk.shape[0]  # Number of observations in regime k

            # Storage for bootstrap iterations
            vol_samples = np.empty((n_boot, self.D))
            pair_keys = [(i, j) for i, j in combinations(range(self.D), 2)]
            corr_samples = np.empty((n_boot, len(pair_keys)))

            for b in range(n_boot):
                # Sample n_k observations with replacement
                idx = rng.choice(nk, size=nk, replace=True)
                Yb = Yk[idx]
                vol_samples[b] = np.std(Yb, axis=0, ddof=1)  # Sample volatility
                C = self._safe_corrcoef(Yb)  # Sample correlation matrix
                for p_idx, (i, j) in enumerate(pair_keys):
                    corr_samples[b, p_idx] = C[i, j]

            lo = alpha / 2 * 100  # Lower bound percentile
            hi = (1 - alpha / 2) * 100  # Upper bound percentile

            # Volatility CIs
            vol_ci = {}
            for d in range(self.D):
                point = np.std(Yk, axis=0, ddof=1)[d]
                vol_ci[self.asset_names[d]] = {
                    "point": point,
                    "ci_low": np.percentile(vol_samples[:, d], lo),
                    "ci_high": np.percentile(vol_samples[:, d], hi),
                }

            # Correlation CIs
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

    def get_bootstrap_ci_df(self, n_boot: int = 5000, alpha: float = 0.05, seed: int = 42, annualise: float = np.sqrt(12)):
        """
        Returns bootstrap results as a DataFrame for clean final display.
        """
        boot_data = self.generate_bootstrap_ci(n_boot=n_boot, alpha=alpha, seed=seed)
        rows = []

        for k, metrics in boot_data.items():
            # 1. Volatilities
            for asset, ci in metrics["vol_ci"].items():
                rows.append({
                    "Regime": k,
                    "Type": "Volatility",
                    "Label": asset,
                    "Value": ci["point"] * annualise,
                    "Lower_Bound": ci["ci_low"] * annualise,
                    "Upper_Bound": ci["ci_high"] * annualise
                })

            # 2. Correlations
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
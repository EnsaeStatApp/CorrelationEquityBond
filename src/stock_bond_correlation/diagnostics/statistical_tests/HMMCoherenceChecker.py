import pandas as pd
import numpy as np
from src.stock_bond_correlation.diagnostics.statistical_tests.EmissionStatisticalAnalyzer import EmissionStatisticalAnalyzer
from src.stock_bond_correlation.diagnostics.statistical_tests.TransitionStatisticalAnalyzer import TransitionStatisticalAnalyzer
from src.stock_bond_correlation.models.base import RegimeDetector
from typing import List, Dict, Tuple
from IPython.display import display


class HMMCoherenceChecker:
    """
    Control center to validate HMM fit robustness.
    Automates separability, numerical stability, and persistence tests.
    """
    def __init__(self, detector: RegimeDetector, log_returns: np.ndarray, asset_names: List[str] = None, covariate_names: List[str] = None,
                 min_duration: float = 3.0, min_freq: float = 0.05, sig_threshold: float = 0.5, sig_max_cond_number: float = 1e8,
                 min_confidence: float = 0.7, weights: Dict[str, float] = None):
        """
        Parameters:
        -----------
        detector : RegimeDetector
            Instance of RegimeDetector (Market or Macro).
        log_returns : np.ndarray
            Financial returns (T, D) for emission tests.
        asset_names : List[str]
            Names of the assets.
        covariate_names : List[str]
            Names of macro variables (if IOHMM).
        min_duration : float
            Minimum accepted average duration (e.g., 3 months).
        min_freq : float
            Minimum long-term frequency (e.g., 5%).
        sig_threshold : float
            Minimum ratio of significant separability tests.
        sig_max_cond_number : float
            Maximum acceptable condition number for the Fisher Information matrix.
        min_confidence : float
            Minimum average confidence index (1-Entropy).
        weights : Dict[str, float]
            Weights for the global score (sum must equal 1).
        """
        self.detector = detector
        self.log_returns = log_returns
        self.asset_names = asset_names
        self.covariate_names = covariate_names

        # Validation thresholds
        self.min_duration = min_duration
        self.min_freq = min_freq
        self.sig_threshold = sig_threshold
        self.sig_max_cond_number = sig_max_cond_number
        self.min_confidence = min_confidence

        # Global score weights
        self.weights = weights or {
            "separability": 0.3,
            "stability": 0.2,
            "persistence": 0.2,
            "confidence": 0.3
        }

    def _refresh_analyzers(self):
        """
        Instantiates analyzers with the detector's current (post-fit) state.
        """
        if not self.detector.is_fitted:
            raise ValueError("The detector must be fitted before using the checker.")

        # Retrieve Viterbi states for the emission analyzer
        states = self.detector.viterbi_states

        # Consistency resolution for log_returns and asset_names:
        # If asset_names is provided explicitly, it defines the analysis dimension.
        # We truncate log_returns to the first N columns (N = len(asset_names)) to avoid
        # analyzing macro/extra columns that might have been concatenated in data_to_fit.
        lr_full = np.array(self.log_returns)
        if lr_full.ndim == 1:
            lr_full = lr_full[:, None]

        if self.asset_names is not None and len(self.asset_names) > 0:
            # asset_names provided -> defines the number of financial columns to analyze
            n_assets = len(self.asset_names)
            lr = lr_full[:, :n_assets]          # Truncate excess columns
            asset_names_safe = list(self.asset_names)
        else:
            # asset_names absent -> take all columns and generate generic names
            lr = lr_full
            asset_names_safe = [f"Asset_{i+1}" for i in range(lr.shape[1])]

        stat_ana = EmissionStatisticalAnalyzer(
            log_returns=lr,
            states=states,
            asset_names=asset_names_safe
        )
        trans_ana = TransitionStatisticalAnalyzer(detector=self.detector)

        return stat_ana, trans_ana

    def check_separability(self):
        """
        Checks if regimes are statistically distinct (Vol & Correl) using EmissionStatisticalAnalyzer.
        """
        stat_ana, _ = self._refresh_analyzers()
        levene = stat_ana.test_pairwise_levene()
        fisher = stat_ana.test_fisher_z()

        total = len(levene) + len(fisher)
        sigs = levene['significant'].sum() + fisher['significant'].sum()
        ratio = sigs / total if total > 0 else 0

        return {"ratio": ratio, "is_ok": ratio >= self.sig_threshold}

    def check_stability(self, ref_idx: int = 0):
        """
        Analyzes the Maximum Likelihood shape (Hessian) via TransitionStatisticalAnalyzer.

        Parameter: 
        ----------
        ref_idx : int
            Reference index for calculating uncertainty in TransitionStatisticalAnalyzer.
        """
        _, trans_ana = self._refresh_analyzers()
        # Calculate numerical inference
        _, diag = trans_ana.compute_transition_inference(
            covariate_names=self.covariate_names,
            ref_idx=ref_idx
        )

        eigvals = diag['eigvals']
        cond_num = diag['cond_num']

        # A model is stable if Fisher eigenvalues are > 0 and the Condition Number is reasonable
        is_stable = (np.all(eigvals > 0)) and (cond_num < self.sig_max_cond_number)

        return {"cond_num": cond_num, "is_stable": is_stable, "min_eig": np.min(eigvals)}


    def report_emission_stats(self, n_boot: int = 5000, alpha: float = 0.05, annualise: float = np.sqrt(12)):
        """
        Displays full emission diagnostics: 
        Volatilities and Correlations with Confidence Intervals (Bootstrap).

        Parameters:
        -----------
        n_boot : int
            Number of bootstrap iterations.
        alpha : float
            Significance threshold.
        annualise : float
            Annualization factor.
        """
        stat_ana, _ = self._refresh_analyzers()
        
        print("\n" + "="*70)
        print("COMPLETE EMISSIONS DIAGNOSTIC (BOOTSTRAP & TESTS)")
        print("="*70)
        
        # 1. Retrieve global Bootstrap DataFrame
        boot_df = stat_ana.get_bootstrap_ci_df(n_boot=n_boot, alpha=alpha, annualise=annualise)
        
        # Clean display of volatilities
        print("\n1. ANNUAL VOLATILITIES WITH CONFIDENCE INTERVALS:")
        vols_boot = boot_df.xs('Volatility', level='Type')
        display(vols_boot.round(4))

        # Clean display of correlations
        print("\n2. CORRELATIONS WITH CONFIDENCE INTERVALS:")
        corrs_boot = boot_df.xs('Correlation', level='Type')
        display(corrs_boot.round(4))

        # Display significance test results
        print("\n3. SEPARABILITY TESTS (ALL PAIRS):")
        
        levene = stat_ana.test_pairwise_levene(alpha=alpha)
        fisher = stat_ana.test_fisher_z(alpha=alpha)
        
        # Concatenate important tests
        print("- Volatility Significance (Levene):")
        display(levene.sort_values("p_value").round(4))
        
        print("\n- Correlation Significance (Fisher-Z):")
        display(fisher.sort_values("p_value").round(4))

    def report_transition_stats(self, ref_idx: int = 0, inputs: np.ndarray = None):
        """
        Displays full transition diagnostics:
        Weight inference (p-values), transition matrices, and durations.

        Parameters:
        -----------
        ref_idx : int
            Reference index for Hessian-based uncertainty calculation.
        inputs : np.ndarray
            Covariates that may impact transitions.
        """
        _, trans_ana = self._refresh_analyzers()

        print("\n" + "="*70)
        print("TRANSITION DIAGNOSTIC (MODEL DYNAMICS)")
        print("="*70)

        # 1. Statistical inference (Hessian / Fisher Info)
        print(f"\n1. PARAMETER SIGNIFICANCE (Ref: Regime {ref_idx}):")
        inf_df, _ = trans_ana.compute_transition_inference(
            covariate_names=self.covariate_names,
            ref_idx=ref_idx
        )
        display(inf_df.round(4))

        # 2. Average dynamics
        print("\n2. AVERAGE DYNAMICS AND PERSISTENCE:")
        avg_A, durations, stationary = trans_ana.get_average_transition_dynamics(inputs=inputs)

        dyn_df = pd.DataFrame({
            "Average_Duration (months)": durations,
            "Stationary_Prob": stationary,
            "Stay_Prob": np.diag(avg_A)
        })
        display(dyn_df.round(4))

    def compute_coherence_score(self, inputs: np.ndarray = None):
        """
        Calculates a weighted score from 0 to 1 reflecting the overall fit quality.
        """
        sep = self.check_separability()
        stab = self.check_stability()

        _, trans_ana = self._refresh_analyzers()
        _, durations, stationary = trans_ana.get_average_transition_dynamics(inputs=inputs)

        # Confidence via entropy (method inherited from BaseRegimeDetector)
        avg_conf = np.mean(self.detector.compute_confidence_index(series_index=0))

        # --- Sub-score normalization ---
        s_sep = sep['ratio']

        # Stability (logarithmic scale for Condition Number)
        log_cond = np.log10(max(stab['cond_num'], 1.0))
        s_stab = max(0, 1 - (log_cond / 12.0))
        if not stab['is_stable']: s_stab *= 0.1 # Critical penalty

        # Persistence (Duration and Frequency)
        s_dur = min(durations.min() / self.min_duration, 1.0)
        s_freq = min(stationary.min() / self.min_freq, 1.0)
        s_pers = (s_dur + s_freq) / 2

        s_conf = avg_conf

        score = (
            self.weights["separability"] * s_sep +
            self.weights["stability"] * s_stab +
            self.weights["persistence"] * s_pers +
            self.weights["confidence"] * s_conf
        )

        details = {
            "sep_ratio": s_sep,
            "is_stable": stab['is_stable'],
            "min_dur": durations.min(),
            "min_freq": stationary.min(),
            "avg_conf": avg_conf,
            "durations": durations,
            "stationary": stationary
        }

        return score, details

    def final_verdict(self, inputs: np.ndarray = None):
        """
        Displays a complete report and decides on model validity.
        """
        score, d = self.compute_coherence_score(inputs=inputs)

        # Go/No-Go Criteria
        c_sep = d["sep_ratio"] >= self.sig_threshold
        c_stab = d["is_stable"]
        c_pers = (d["min_dur"] >= self.min_duration) and (d["min_freq"] >= self.min_freq)
        c_conf = d["avg_conf"] >= self.min_confidence

        print("\n" + "="*50)
        print(f"HMM COHERENCE REPORT : {score:.2%}")
        print("="*50)
        print(f"1. SEPARABILITY : {d['sep_ratio']:.1%} (Threshold: {self.sig_threshold:.0%}) -> {'OK' if c_sep else 'WEAK'}")
        print(f"2. STABILITY    : {'VALID' if c_stab else 'FAILED (Singularity)'}")
        print(f"3. PERSISTENCE  : {d['min_dur']:.1f} months (min {self.min_duration}) -> {'OK' if c_pers else 'TOO SHORT'}")
        print(f"4. CONFIDENCE   : {d['avg_conf']:.1%} (Entropy) -> {'SOLID' if c_conf else 'HESITANT'}")
        print("-" * 50)

        df_dyn = pd.DataFrame({"Duration": d['durations'], "Freq": d['stationary']})
        print("REGIME DYNAMICS:")
        print(df_dyn.round(3).to_string())

        if c_sep and c_stab and c_pers and c_conf:
            print("\nVERDICT: MODEL VALIDATED (Robust)")
            return score
        else:
            print("\nVERDICT: MODEL REJECTED (Instable or non-significant)")
            return 0.0 # Return 0 for rejected models

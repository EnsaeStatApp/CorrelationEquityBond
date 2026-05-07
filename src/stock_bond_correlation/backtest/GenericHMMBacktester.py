import pandas as pd
import numpy as np
from tqdm import tqdm
import copy
from typing import List, Union, Dict
from joblib import Parallel, delayed
from RegimeDetector.HMMCoherenceChecker import HMMCoherenceChecker

# =============================================================================
# 1. TOURNAMENT FUNCTION (EXTERNAL)
# =============================================================================
def _evaluate_model_candidate(tester: 'GenericHMMBacktester', k: int, method_name: str, seed: int, 
                              observations_train: np.ndarray, inputs_train: np.ndarray) -> tuple:
    """
    Evaluates a single HMM model candidate for the tournament selection process.

    This function clones the detector, fits it on the training data using a specific seed 
    and initialization method, and returns a performance score (Log-Likelihood or Coherence Score).

    Args:
        - tester (GenericHMMBacktester): The backtester instance providing cloning and checking utilities.
        - k (int): Number of regimes (states) to use in the HMM.
        - method_name (str): Initialization method for the HMM (e.g., "random", "kmeans").
        - seed (int): Random seed for reproducibility of the initialization.
        - observations_train (np.ndarray): Scaled training observations matrix of shape (n_samples, n_features).
        - inputs_train (np.ndarray): Normalized input features (macro drivers) of shape (n_samples, n_inputs).

    Returns:
        tuple: A tuple containing (score, detector). 
            - score (float): The evaluation metric (log-likelihood or coherence score). Returns -inf if failed.
            - detector (object): The fitted HMM detector instance or None if the model was unstable.
    """
    det = tester._clone_detector(k, seed)
    try:
        det.fit(observations=[observations_train],
                inputs=[inputs_train] if tester.input_vars else None,
                warm_start=True,
                initialize=True,
                init_method=method_name,
                num_iters=200)
        D = observations_train.shape[1]
        
        # If single asset, use Log-Likelihood as performance metric
        if D == 1:
            return det.model.log_likelihood(observations_train, inputs=inputs_train), det
        else:
            # Check if all regimes have at least 2 observations assigned
            if np.min(np.bincount(det.viterbi_states, minlength=k)) >= 2:
                checker = tester._create_checker(det, observations_train)
                score, details = checker.compute_coherence_score(inputs=inputs_train)
                # Ensure the model is stable and respects the minimum separation threshold
                if details["is_stable"] and details["sep_ratio"] >= tester.sig_threshold:
                    return score, det
    except:
        pass
    return -np.inf, None


# =============================================================================
# 2. BACKTESTER CLASS WITH SAFETY FEATURES AND INTEGRATED AUDIT
# =============================================================================
class GenericHMMBacktester:
    def __init__(self, detector: object, data_df: pd.DataFrame, obs_vars: List[str], asset_vars: List[str],
             input_vars: List[str] = None, refit_freq: int = 1, reset_total_freq: int = 12,
             scaling_factor: float = 100.0, sig_threshold: float = 0.5, n_jobs: int = -1,
             min_nk_threshold: float = 10.0, **checker_params):
        """
        Initializes the HMM Backtester for causal regime detection and portfolio analysis.

        Args:
            - detector (object): A template HMM detector object following the fit/predict API.
            - data_df (pd.DataFrame): The historical dataset containing observations, assets, and inputs.
            - obs_vars (List[str]): Column names of the variables used to train the HMM.
            - asset_vars (List[str]): Column names of the assets for which risk metrics are computed.
            - input_vars (List[str], optional): Column names of exogenous macro drivers. Defaults to None.
            - refit_freq (int): Frequency (in steps) at which the model parameters are updated. Defaults to 1.
            - reset_total_freq (int): Frequency at which a full re-initialization (tournament) occurs. Defaults to 12.
            - scaling_factor (float): Multiplier for observations to improve numerical stability. Defaults to 100.0.
            - sig_threshold (float): Minimum separation ratio for regime coherence. Defaults to 0.5.
            - n_jobs (int): Number of CPU cores for parallel processing. -1 uses all cores. Defaults to -1.
            - min_nk_threshold (float): Minimum effective sample size per regime to avoid "zombie" states. Defaults to 10.0.
            - **checker_params: Additional arguments passed to the HMMCoherenceChecker.
        """

        self.df = data_df
        self.obs_vars, self.asset_vars = obs_vars, asset_vars
        self.input_vars = input_vars or []
        self.template = detector
        self.refit_freq, self.reset_total_freq = refit_freq, reset_total_freq
        self.scaling, self.n_jobs = scaling_factor, n_jobs
        self.sig_threshold = sig_threshold
        self.min_nk = min_nk_threshold
        self.checker_params = checker_params
        self.history_detectors = {}

        self.current_detector = None
        self._last_sigmas, self._last_means, self._last_nk = None, None, None

        # Result storage
        self.results = {}
        self.regime_probs = {}
        self.regime_sigmas = {}
        self.regime_means = {}
        self.regime_nk = {}
        self.history_logs = []

    def _clone_detector(self, n_regimes: int, seed: int) -> object:
        """
        Creates a fresh copy of the template detector with specific parameters.

        Args:
            - n_regimes (int): Number of hidden states for the new instance.
            - seed (int): Random state for initialization.

        Returns:
            object: A new, unfitted HMM detector instance.
        """
        new_det = copy.deepcopy(self.template)
        t_kwargs = getattr(self.template, 'transition_kwargs', None)
        obs_type = getattr(self.template, 'observations_type', 'gaussian')
        new_det.__init__(n_regimes=n_regimes, n_dim=len(self.obs_vars),
                         n_input=len(self.input_vars), random_state=seed,
                         transition_kwargs=t_kwargs, observations_type=obs_type)
        return new_det

    def _create_checker(self, detector: object, observations_scaled: np.ndarray) -> object:
        """
        Instantiates the coherence checker to validate the economic meaning of regimes.

        Args:
            - detector (object): The fitted HMM model.
            - observations_scaled (np.ndarray): The data used for validation.

        Returns:
            HMMCoherenceChecker: An object to compute stability and separation metrics.
        """
        return HMMCoherenceChecker(detector, observations_scaled, self.asset_vars,
                                   sig_threshold=self.sig_threshold, **self.checker_params)

    def _do_reset(self, obs_tr: np.ndarray, in_tr: np.ndarray) -> tuple:
        """
        Performs a parallel tournament of 10 random initializations to find the most stable model.

        Args:
            - obs_tr (np.ndarray): Training observations.
            - in_tr (np.ndarray): Training inputs.

        Returns:
            tuple: (best_detector, method_name) or (None, "FAILED") if no candidate is stable.
        """
        seeds = np.linspace(0, 10000, 10).astype(int)
        results = Parallel(n_jobs=self.n_jobs)(
            delayed(_evaluate_model_candidate)(self, self.template.K, "random", s, obs_tr, in_tr)
            for s in seeds
        )
        candidates = [r for r in results if r[0] != -np.inf]
        if candidates:
            best = max(candidates, key=lambda x: x[0])
            return best[1], "Reset-Random"
        else:
            return None, "FAILED"

    def run(self, start_date: Union[str, pd.Timestamp], end_date: Union[str, pd.Timestamp] = "2025-12-01", start_training_date : pd.Timestamp = "1974-01-01"):
        """
        Executes the causal backtest loop over the specified date range.

        This method handles the data windowing, model fitting (warm-starts or resets), 
        and generates monthly regime audits including macro-driver influences.

        Args:
            - start_date (Union[str, pd.Timestamp]): The date to begin the backtest.
            - end_date (Union[str, pd.Timestamp]): The date to end the backtest. Defaults to "2025-12-01".
            - start_training_date (pd.Timestamp): Start of the expanding window for training. Defaults to "1974-01-01".

        Raises:
            ValueError: If no stable model can be found for the initial starting date.
        """
        test_indices = self.df.index[(self.df.index >= start_date) & (self.df.index <= end_date)]

        for i, current_date in enumerate(tqdm(test_indices)):
            # --- A. CAUSAL DATA PREPARATION ---
            # Define training window up to (but excluding) current_date
            train_mask = (self.df.index >= start_training_date) & (self.df.index < current_date)
            obs_tr = self.df.loc[train_mask, self.obs_vars].ffill().bfill().values.astype(float) * self.scaling
            # Evaluation data including current_date for prediction
            obs_f = self.df.loc[self.df.index <= current_date, self.obs_vars].ffill().bfill().values * self.scaling

            in_tr, in_f = None, None
            if self.input_vars:
                raw_in_tr = self.df.loc[train_mask, self.input_vars].ffill().bfill().values.astype(float)
                raw_in_f = self.df.loc[self.df.index <= current_date, self.input_vars].ffill().bfill().values.astype(float)
                # Causal normalization of inputs
                in_mu = np.mean(raw_in_tr, axis=0)
                in_std = np.std(raw_in_tr, axis=0) + 1e-9
                in_tr = (raw_in_tr - in_mu) / in_std
                in_f = (raw_in_f - in_mu) / in_std

            # --- B. TRAINING PHASE ---
            method = "Predict-Only"
            if i % self.refit_freq == 0 or self.current_detector is None:
                force_reset = (i > 0 and i % self.reset_total_freq == 0)
                old_detector = copy.deepcopy(self.current_detector) if self.current_detector is not None else None

                try:
                    if force_reset or self.current_detector is None:
                        # Full tournament reset
                        new_det, m = self._do_reset(obs_tr, in_tr)
                        
                        # Fallback for the first iteration: decrease threshold if no model is found
                        if new_det is None and i == 0:
                            print(f"\n[FIRST FIT] Failed with threshold={self.sig_threshold:.2f}. Emergency attempt with {self.sig_threshold - 0.1:.2f}...")
                            orig_sig = self.sig_threshold
                            self.sig_threshold -= 0.1
                            new_det, m = self._do_reset(obs_tr, in_tr)
                            self.sig_threshold = orig_sig

                        if new_det is not None:
                            self.current_detector, method = new_det, m
                        else:
                            if old_detector is not None:
                                self.current_detector = old_detector
                                method = "Keep-Previous (Tournament Failed)"
                            else:
                                raise ValueError(f"Could not find an initial model at {current_date.date()}")
                    else:
                        # Warm-start refit
                        self.current_detector.fit(observations=[obs_tr], inputs=[in_tr] if in_tr is not None else None,
                                                  warm_start=False, initialize=False, num_iters=50)
                        method = "Warm-Start"
                        
                        # Zombie Protection: if a regime becomes too small, force a reset
                        nk_check = np.sum(self.current_detector.regime_probabilities(series_index=0), axis=0)
                        if np.min(nk_check) < self.min_nk:
                            new_det, m = self._do_reset(obs_tr, in_tr)
                            if new_det is not None:
                                self.current_detector, method = new_det, m
                            else:
                                self.current_detector = old_detector
                                method = "Keep-Previous (Zombie Protection)"
                except Exception as e:
                    if old_detector is not None:
                        self.current_detector = old_detector
                        method = "Keep-Previous (Error Recovery)"
                    else:
                        raise e

                # Store fitted parameters (unscaled)
                self.history_detectors[current_date] = copy.deepcopy(self.current_detector)
                nk = np.sum(self.current_detector.regime_probabilities(series_index=0), axis=0)
                m_scaled = self.current_detector.regime_means()
                s_scaled = self.current_detector.regime_covariances()
                self._last_means = [m / self.scaling for m in m_scaled]
                self._last_sigmas = [s / (self.scaling**2) for s in s_scaled]
                self._last_nk = nk

            # --- C. PREDICTION AND AUDIT ---
            self.regime_means[current_date] = self._last_means
            self.regime_sigmas[current_date] = self._last_sigmas
            self.regime_nk[current_date] = self._last_nk

            try:
                # OOS prediction for the current month
                proba_now = self.current_detector.predict_probabilities(obs_f, inputs=in_f, oos_start=len(obs_f)-1)[0]
                proba_now = (proba_now + 1e-12) / np.sum(proba_now + 1e-12)
            except:
                proba_now = np.ones(self.template.K)/self.template.K

            self.regime_probs[current_date] = proba_now
            # Weighted expected covariance matrix
            self.results[current_date] = sum(proba_now[k] * self.regime_sigmas[current_date][k] for k in range(self.current_detector.K))

            # --- ROLLING STATS AND REALIZED RETURNS (CAUSAL) ---
            # Current month's returns
            realized_rets = self.df.loc[current_date, self.asset_vars].values
            # 12-month rolling stats (causal: ending at current_date)
            hist_12m = self.df.loc[self.df.index <= current_date, self.asset_vars].tail(12)
            if len(hist_12m) >= 2:
                rolling_cov = hist_12m.cov() * 12
                rolling_vols = np.sqrt(np.diag(rolling_cov))
                rolling_rho = rolling_cov.iloc[0, 1] / (rolling_vols[0] * rolling_vols[1] + 1e-15)
            else:
                rolling_vols, rolling_rho = [0, 0], 0

            # --- D. DETAILED AUDIT DISPLAY ---
            print(f"\n" + "━"*115)
            print(f"▶ DATE : {current_date.date()} | MODE : {method}")
            print(f"  [MONTHLY REALIZED RETS] {self.asset_vars[0]}: {realized_rets[0]:>+7.2%} | {self.asset_vars[1]}: {realized_rets[1]:>+7.2%}")
            print(f"  [ROLLING 12M STATS]     V_SP: {rolling_vols[0]:>5.1%} | V_BN: {rolling_vols[1]:>5.1%} | RHO: {rolling_rho:>5.2f}")

            if self.input_vars:
                x_now = in_f[-1]
                macro_str = " | ".join([f"{name}: {val:+.2f}" for name, val in zip(self.input_vars, x_now)])
                print(f"  [INPUTS MACRO X_t]      {macro_str}")
                # Compute current transition matrix based on macro inputs
                log_Ps, Ws = self.current_detector.get_transition_params()
                log_T_t = log_Ps + Ws @ x_now
                e_x = np.exp(log_T_t - np.max(log_T_t, axis=1, keepdims=True))
                T_t = e_x / e_x.sum(axis=1, keepdims=True)

            print(f"  {'-'*110}")
            w_headers = " | ".join([f"W_{m[:8]:<10}" for m in self.input_vars])
            header = f"  {'REGIME':<8} | {'PROBA':<6} | {'RHO':<5} | {'V_SP':<6} | {'V_BN':<6} | {w_headers}"
            print(header)
            print(f"  {'-'*110}")

            for k in range(self.template.K):
                sig_k = self.regime_sigmas[current_date][k]
                vols = [np.sqrt(sig_k[j,j]*12) for j in range(2)]
                rho = sig_k[0,1] / (vols[0]*vols[1]/12 + 1e-15)
                # Display macro driver weights for each regime transition
                weights_k = " | ".join([f"{Ws[k, m]:>+12.2f}" for m in range(len(self.input_vars))]) if self.input_vars else ""
                row = f"  Reg {k:<4} | {proba_now[k]:>5.1%} | {rho:>5.2f} | {vols[0]:>5.1%} | {vols[1]:>5.1%} | {weights_k}"
                print(row)
            print("━"*115)

    def get_risk_df(self) -> pd.DataFrame:
        """
        Compiles the estimated volatilities for each asset into a structured DataFrame.

        Returns:
            pd.DataFrame: A time-series of annualized volatilities derived from the regime-weighted covariances.
        """
        dates = sorted(self.results.keys())
        cols = [f"Vol_{a}" for a in self.asset_vars]
        data = [[np.sqrt(self.results[d][i,i]*12) for i in range(len(self.asset_vars))] for d in dates]
        return pd.DataFrame(data, index=dates, columns=cols)
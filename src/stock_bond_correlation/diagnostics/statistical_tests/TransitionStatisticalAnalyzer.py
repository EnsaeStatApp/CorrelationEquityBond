import numpy as np
import pandas as pd
from scipy import stats as sp_stats
from typing import List
from src.stock_bond_correlation.models.base_detector import BaseRegimeDetector

class TransitionStatisticalAnalyzer:
    """
    Class for performing descriptive statistical tests on a detector's transition parameters.

    Features:
    - Extracts transition parameters for each regime (intercepts + macro weights if present).
    - Computes Standard Errors (SE) for these parameters by numerically approximating 
      the model's Fisher Information matrix.
    - Computes the eigenvalues of the Hessian at the optimal point (transition parameters 
      only) to determine the "shape" of the MLE peak (curvature/stability).
    - Returns transition matrices associated with covariates (time-varying) or constants.
    - Calculates the expected average duration of each regime.
    """

    def __init__(self, detector: BaseRegimeDetector):
        """
        Initializes the analyzer with a regime detector.

        Args:
            detector (BaseRegimeDetector): A general regime detector instance.
        """
        self.detector = detector
        self.K = detector.K
        self.M = detector.M

    def compute_transition_inference(self, covariate_names: List[str] = None, series_index: int = 0, 
                                     alpha: float = 0.05, ref_idx: int = 0, epsilon: float = 1e-4):
        """
        Computes uncertainties for transition probabilities (intercepts and macro weights).

        The Fisher Information matrix is computed by approximating the Log-Likelihood 
        Hessian using finite differences.

        Args:
            covariate_names (List[str], optional): Names of exogenous covariates.
            series_index (int): Index of the data sequence to analyze.
            alpha (float): Significance threshold for p-values.
            ref_idx (int): Reference regime index for parameter centering/identification.
            epsilon (float): Step size for finite difference approximation.

        Returns:
            tuple: (DataFrame of coefficients/p-values, dictionary of spectral diagnostics).
        """

        # 1. Backup original data to restore the model state after inference
        data = np.array(self.detector.fit_observations[series_index])
        if self.detector.fit_input is not None:
            inp = np.array(self.detector.fit_input[series_index])
        else:
            inp = None

        # 2. Backup current transition parameters
        log_Ps_orig, Ws_orig = self.detector.get_transition_params()

        # 3. Parameter centering: reduce parameters by relative comparison to a reference regime
        # Center intercepts (log_Ps)
        reference_column = log_Ps_orig[:, [ref_idx]] 
        log_Ps_centered = log_Ps_orig - reference_column
        log_Ps_r = np.delete(log_Ps_centered, ref_idx, axis=1)

        # Center macro weights (Ws)
        reference_row = Ws_orig[[ref_idx], :] 
        Ws_centered = Ws_orig - reference_row
        Ws_r = np.delete(Ws_centered, ref_idx, axis=0)

        n_logPs = self.K * (self.K - 1)
        p_red = n_logPs + (self.K - 1) * self.M # Reduced parameter count
        # Flatten parameters into a single theta vector for Hessian calculation
        params_opt = np.concatenate([log_Ps_r.flatten(), Ws_r.flatten()]) 

        # 4. Helper function to compute LL from a flat parameter vector
        def get_ll(params_flat: np.ndarray):
            """
            Maps a flat parameter vector back to log_Ps/Ws and computes the Log-Likelihood.

            Args:
                params_flat (np.ndarray): Flattened parameter vector.
            """
            lP_r = params_flat[:n_logPs].reshape(self.K, self.K - 1)
            W_r = params_flat[n_logPs:].reshape(self.K - 1, self.M)

            # Reconstruct full matrices by re-inserting the reference column/row (zeros)
            lP_f = np.concatenate([lP_r[:, :ref_idx], np.zeros((self.K, 1)), lP_r[:, ref_idx:]], axis=1)
            W_f = np.concatenate([W_r[:ref_idx, :], np.zeros((1, self.M)), W_r[ref_idx:, :]], axis=0)

            # Update parameters and compute LL via the detector wrapper
            self.detector.update_transition_params(lP_f, W_f) 
            return self.detector.compute_ll([data], [inp]) 

        # 5. Compute the Hessian via second-order centered finite differences
        ll_base = get_ll(params_opt)
        H = np.zeros((p_red, p_red))
        
        # Diagonal elements
        for i in range(p_red): 
            p_plus = params_opt.copy()
            p_plus[i] += epsilon
            p_minus = params_opt.copy()
            p_minus[i] -= epsilon
            ll_p = get_ll(p_plus)
            ll_m = get_ll(p_minus)

            H[i, i] = (ll_p - 2 * ll_base + ll_m) / (epsilon ** 2)

        # Cross-partial derivatives
        for i in range(p_red): 
            for j in range(i + 1, p_red): 
                p_pp = params_opt.copy()
                p_pp[i] += epsilon
                p_pp[j] += epsilon
                p_mm = params_opt.copy()
                p_mm[i] -= epsilon
                p_mm[j] -= epsilon
                p_pm = params_opt.copy()
                p_pm[i] += epsilon
                p_pm[j] -= epsilon
                p_mp = params_opt.copy()
                p_mp[i] -= epsilon
                p_mp[j] += epsilon
                val = (get_ll(p_pp) - get_ll(p_pm) - get_ll(p_mp) + get_ll(p_mm)) / (4 * epsilon ** 2)
                H[i, j], H[j, i] = val, val

        # Restore initial parameters to the model
        self.detector.update_transition_params(log_Ps_orig, Ws_orig)  

        # 6. Spectral analysis (for numerical stability) and Matrix Inversion
        I_obs = - H # Observed Fisher Information
        eigvals = np.linalg.eigvalsh(I_obs)
        cond_num = eigvals.max() / (np.abs(eigvals.min()) + 1e-12)

        try: 
            I_inv = np.linalg.inv(I_obs) # Variance-Covariance Matrix
        except np.linalg.LinAlgError:
            I_inv = np.linalg.pinv(I_obs) # Use pseudo-inverse if non-invertible

        # 7. Standard Error calculation and Significance Testing
        se = np.sqrt(np.maximum(np.diag(I_inv), 1e-12))
        se_lPs_r = se[:n_logPs].reshape(self.K, self.K - 1)
        se_Ws_r = se[n_logPs:].reshape(self.K - 1, self.M)

        # Final preparation of results
        rows = []
        regimes_dest = [r for r in range(self.K) if r != ref_idx]
        cov_names = covariate_names or [f"X{m}" for m in range(self.M)]

        # A. Transition Intercepts (Base Probabilities)
        for i, k_origin in enumerate(range(self.K)):
            for j, k_dest in enumerate(regimes_dest):
                coef = log_Ps_r[i, j]
                stderr = se_lPs_r[i, j]
                z = coef / stderr
                pval = 2 * sp_stats.norm.sf(np.abs(z))
                rows.append({
                    'Type': 'Intercept',
                    'Transition': f"R{k_origin} -> R{k_dest}",
                    'Variable': 'Base_Prob',
                    'Coef': coef, 'Std_Err': stderr, 'P_Value': pval,
                    'Significant': pval < alpha
                })

        # B. Macro Weights (Impact of drivers on transitions)
        for i, k_dest in enumerate(regimes_dest):
            for m in range(self.M):
                coef = Ws_r[i, m]
                stderr = se_Ws_r[i, m]
                z = coef / stderr
                pval = 2 * sp_stats.norm.sf(np.abs(z))
                rows.append({
                    'Type': 'Macro_Weight',
                    'Transition': f"To R{k_dest} (vs R{ref_idx})",
                    'Variable': cov_names[m],
                    'Coef': coef, 'Std_Err': stderr, 'P_Value': pval,
                    'Significant': pval < alpha
                })

        return pd.DataFrame(rows), {'eigvals': eigvals, 'cond_num': cond_num}


    def get_transition_matrix(self, inputs: np.ndarray = None, mode: str = "average"):
        """
        Calculates the transition matrix based on the specified mode.

        Args:
            inputs (np.ndarray, optional): Exogenous inputs for time-varying matrices.
            mode (str): Aggregation method ('average', 'last', or 'all').

        Returns:
            np.ndarray: Transition matrix (K, K) or array of matrices (T, K, K).
        """
        matrices = self.detector.get_transition_matrices(inputs)

        # ADJUSTMENT: If the detector returns a single matrix (K, K) instead of a stack (T, K, K)
        if matrices.ndim == 2:
            return matrices  

        # If a temporal stack (T, K, K), apply aggregation logic
        if mode == "average":
            return np.mean(matrices, axis=0)
        elif mode == "last":
            return matrices[-1]
        elif mode == "all":
            return matrices
        else:
            raise ValueError("Mode must be 'average', 'last', or 'all'.")

    def get_expected_durations(self, transition_matrix: np.ndarray):
        """
        Calculates the expected duration in each regime from a transition matrix.
        
        Formula: 1 / (1 - P(stay)) (Geometric distribution parameter).

        Args:
            transition_matrix (np.ndarray): Stochastic transition matrix (K x K).

        Returns:
            np.ndarray: Vector of expected durations.
        """
        return 1 / (1 - np.diag(transition_matrix))

    def get_average_transition_dynamics(self, inputs: np.ndarray = None):
        """
        Computes transition dynamics: (Average Matrix, Durations, Long-term Frequencies).

        Args:
            inputs (np.ndarray, optional): Covariates for time-varying models.

        Returns:
            tuple: (avg_A, durations, stationary_distribution).
        """
        # 1. Retrieve the average transition matrix
        avg_A = self.get_transition_matrix(inputs, mode="average")

        # 2. Calculate expected durations
        durations = self.get_expected_durations(avg_A)

        # 3. Calculate the stationary distribution (Long-term frequencies)
        # Find the left eigenvector associated with the eigenvalue 1
        vals, vecs = np.linalg.eig(avg_A.T) 
        stationary = np.real(vecs[:, np.isclose(vals, 1)])
        stationary = stationary[:, 0] / stationary.sum() # Normalization

        return avg_A, durations, stationary
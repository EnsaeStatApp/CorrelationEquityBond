import numpy as np
import pandas as pd
from scipy import stats as sp_stats
from typing import List
from src.stock_bond_correlation.models.base_detector import BaseRegimeDetector

class TransitionStatisticalAnalyzer:
    """
    Class that performs descriptive statistical tests on a detector's transition parameters.

    - Returns the transition parameters for each regime (intercept + weights associated with covariates if present).
    - Returns the standard errors associated with these parameters by numerically calculating the model's Fisher information.
    - Returns the eigenvalues of the Hessian at the optimal point (only on transition parameters for simplicity) 
      to determine the "shape of the MLE peak".
    - Returns transition matrices associated with covariates if present; otherwise, they are constant.
    - Returns the expected duration of each regime linked to a transition matrix.
    """

    def __init__(self, detector: BaseRegimeDetector):
        """
        Parameters:
        -----------
        detector : BaseRegimeDetector
            A general regime detector (assumes detector.fit_data contains log returns).
        """
        self.detector = detector
        self.K = detector.K
        self.M = detector.M

    def compute_transition_inference(self, covariate_names: List[str] = None, series_index: int = 0, alpha: float = 0.05,
                                     ref_idx: int = 0, epsilon: float = 1e-4):
        """
        Calculates uncertainties on covariate weights concerning transition probabilities.

        This is done by calculating the Fisher information, approximating the Hessian of the LL via finite differences.

        Parameters:
        -----------
        covariate_names : List[str], optional
            Names of the covariates.
        series_index : int
            Index of the series to analyze parameters on (reminder: self.detector.fit_data is a list of log returns arrays).
        alpha : float
            Significance threshold.
        ref_idx : int
            Reference regime index used to evaluate significance.
        epsilon : float
            Finite difference step size.
        """

        # 1. Save data because we want the model to return to its original state at the end
        data = np.array(self.detector.fit_observations[series_index])
        if self.detector.fit_input is not None:
            inp = np.array(self.detector.fit_input[series_index])
        else:
            inp = None

        # 2. Save current transition parameters because we want the model to return to its original state at the end
        log_Ps_orig, Ws_orig = self.detector.get_transition_params()

        # 3. Here, we center the parameters around a reference regime
        reference_column = log_Ps_orig[:, [ref_idx]]  # Take the column of the regime in question (using list to keep K*1 matrix)
        log_Ps_centered = log_Ps_orig - reference_column
        log_Ps_r = np.delete(log_Ps_centered, ref_idx, axis=1)

        reference_column = Ws_orig[[ref_idx], :]  # Same approach
        Ws_centered = Ws_orig - reference_column
        Ws_r = np.delete(Ws_centered, ref_idx, axis=0)

        n_logPs = self.K * (self.K - 1)
        p_red = n_logPs + (self.K - 1) * self.M  # Number of reduced parameters compared to the reference regime
        params_opt = np.concatenate([log_Ps_r.flatten(), Ws_r.flatten()])  # To compute the Hessian, flatten parameters into 2 lists then concatenate into a large theta variable

        # 4. Log-Likelihood (LL) function
        def get_ll(params_flat: np.ndarray):
            """
            Takes a parameter vector, reconstructs log_Ps and Ws, and computes the LL.

            Parameter:
            ----------
            params_flat : np.ndarray
                Vector of parameters.
            """
            lP_r = params_flat[:n_logPs].reshape(self.K, self.K - 1)
            W_r = params_flat[n_logPs:].reshape(self.K - 1, self.M)

            lP_f = np.concatenate([lP_r[:, :ref_idx], np.zeros((self.K, 1)), lP_r[:, ref_idx:]], axis=1)
            W_f = np.concatenate([W_r[:ref_idx, :], np.zeros((1, self.M)), W_r[ref_idx:, :]], axis=0)

            self.detector.update_transition_params(lP_f, W_f)  # Update via wrapper

            return self.detector.compute_ll([data], [inp])  # Compute LL via wrapper

        # 5. Hessian calculation via finite differences (second-order centered finite difference)
        ll_base = get_ll(params_opt)
        H = np.zeros((p_red, p_red))
        for i in range(p_red):  # Diagonal
            p_plus = params_opt.copy()
            p_plus[i] += epsilon
            p_minus = params_opt.copy()
            p_minus[i] -= epsilon
            ll_p = get_ll(p_plus)
            ll_m = get_ll(p_minus)

            H[i, i] = (ll_p - 2 * ll_base + ll_m) / (epsilon ** 2)  # Centered difference formula

        for i in range(p_red):  # Cross terms
            for j in range(i + 1, p_red):  # Start at i+1 because H is symmetric
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
                val = (get_ll(p_pp) - get_ll(p_pm) - get_ll(p_mp) + get_ll(p_mm)) / (4 * epsilon ** 2)  # Centered difference formula
                H[i, j], H[j, i] = val, val

        self.detector.update_transition_params(log_Ps_orig, Ws_orig)  # Restore initial parameters

        # 6. Spectral analysis (for numerical stability) and matrix inversion
        I_obs = -H
        eigvals = np.linalg.eigvalsh(I_obs)
        cond_num = eigvals.max() / (np.abs(eigvals.min()) + 1e-12)

        try:  # Attempt standard inversion
            I_inv = np.linalg.inv(I_obs)
        except np.linalg.LinAlgError:  # Fallback to pseudo-inverse
            I_inv = np.linalg.pinv(I_obs)

        # 7. Standard errors calculation
        se = np.sqrt(np.maximum(np.diag(I_inv), 1e-12))
        se_Ws_r = se[n_logPs:].reshape(self.K - 1, self.M)
        se_lPs_r = se[:n_logPs].reshape(self.K, self.K - 1)

        z_Ws = Ws_r / se_Ws_r
        pv_Ws = 2 * sp_stats.norm.sf(np.abs(z_Ws))

        # Prepare final dataframe
        rows = []
        regimes_dest = [r for r in range(self.K) if r != ref_idx]
        cov_names = covariate_names or [f"X{m}" for m in range(self.M)]

        # A. Add Intercepts (log_Ps)
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
                    'Significatif': pval < alpha
                })

        # B. Add Macro Weights (Ws)
        for i, k_dest in enumerate(regimes_dest):
            for m in range(self.M):
                coef = Ws_r[i, m]
                stderr = se_Ws_r[i, m]
                z = coef / stderr
                pval = 2 * sp_stats.norm.sf(np.abs(z))
                rows.append({
                    'Type': 'Macro_Weight',
                    'Transition': f"Vers R{k_dest} (vs R{ref_idx})",
                    'Variable': cov_names[m],
                    'Coef': coef, 'Std_Err': stderr, 'P_Value': pval,
                    'Significatif': pval < alpha
                })

        return pd.DataFrame(rows), {'eigvals': eigvals, 'cond_num': cond_num}

    def get_transition_matrix(self, inputs: np.ndarray = None, mode: str = "average"):
        """
        Calculates the transition matrix.
        """
        matrices = self.detector.get_transition_matrices(inputs)

        # ADJUSTMENT: If the detector returns a single matrix (K, K) instead of (T, K, K)
        if matrices.ndim == 2:
            return matrices  # Already a matrix, nothing to average!

        # If it's a stack (T, K, K), apply the requested logic
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
        Calculates the expected duration in each regime from a matrix A.
        Formula: 1 / (1 - P(stay)) (geometric distribution with parameter transition_matrix[i, i] for regime i).

        Parameters:
        -----------
        transition_matrix : np.ndarray
            Transition matrix.
        """
        return 1 / (1 - np.diag(transition_matrix))

    def get_average_transition_dynamics(self, inputs: np.ndarray = None):
        """
        Returns: (Average Matrix, Durations, Long-Term Frequencies)

        Parameters:
        -----------
        inputs : np.ndarray, optional
            Arrays of covariates impacting transitions.
        """
        # 1. Retrieve the average matrix
        avg_A = self.get_transition_matrix(inputs, mode="average")

        # 2. Calculate expected durations
        durations = self.get_expected_durations(avg_A)

        # 3. Calculate stationary distribution (Long-Term Frequencies)
        vals, vecs = np.linalg.eig(avg_A.T)  # Find the left eigenvector associated with eigenvalue 1
        stationary = np.real(vecs[:, np.isclose(vals, 1)])
        stationary = stationary[:, 0] / stationary.sum()  # Normalization

        return avg_A, durations, stationary
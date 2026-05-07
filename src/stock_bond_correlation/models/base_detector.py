import numpy as np
from src.stock_bond_correlation.models.base import RegimeDetector
import ssm
from typing import List

class BaseRegimeDetector(RegimeDetector):
    """
    Intermediate base class providing common implementations for all SSM-based detectors.
    
    Purpose:
    To centralize shared logic and mathematical computations across different HMM variants, 
    avoiding redundant implementations of probability filtering and risk metric calculations.

    Note:
    Any class inheriting from this is assumed to use the 'ssm' library for its underlying model.
    """

    def conditional_covariance(self, probs: np.ndarray, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Calculates the total conditional covariance using the Law of Total Covariance.

        Formula applied: Total Cov = E[Cov(Y|Z)] + Cov(E[Y|Z])
        Where the first term is the intra-regime variance and the second is the inter-regime variance.

        Args:
            probs (np.ndarray): Array of probabilities of shape (T, K).
            log_returns (np.ndarray, optional): Historical log returns, required for dynamic (EWMA) models.

        Returns:
            np.ndarray: Array of shape (T, D, D) containing the total conditional covariance at each step.
        """
        means = np.array(self.regime_means(log_returns)) # (K, D) or (T, K, D)
        covs = np.array(self.regime_covariances(log_returns)) # (K, D, D) or (T, K, D, D)
        
        T, K = probs.shape
        D = means.shape[-1]
        total_covs = np.zeros((T, D, D))

        for t in range(T):
            p_t = probs[t] # Vector (K,)
            
            # 1. Estimated Mean = sum_k (p_k * mu_kt)
            # Handle both static (means per regime) and dynamic (time-varying means) cases
            m_t = means if means.ndim == 2 else means[t] 
            mu_bar_t = p_t @ m_t 

            # 2. Intra-regime component (Expected Value of Covariances)
            c_t = covs if covs.ndim == 3 else covs[t] 
            intra_t = np.average(c_t, weights=p_t, axis=0) # Weighted average of cov matrices

            # 3. Inter-regime component (Covariance of Expected Values)
            inter_t = np.zeros((D, D))
            for k in range(K):
                diff = m_t[k] - mu_bar_t
                inter_t += p_t[k] * np.outer(diff, diff)

            total_covs[t] = intra_t + inter_t
            
        return total_covs
    

    def regime_correlations(self, log_returns: np.ndarray = None):
        """
        Computes the correlation matrices for each regime k.

        Args:
            log_returns (np.ndarray, optional): Historical log returns, required for dynamic models.

        Returns:
            Union[List[np.ndarray], np.ndarray]: 
                - Static: List of K correlation matrices of shape (D, D).
                - Dynamic: Array of shape (T, K, D, D) of time-varying correlation matrices.
        """
        sigmas = np.array(self.regime_covariances(log_returns))
        
        # STATIC CASE (K, D, D)
        if sigmas.ndim == 3:    
            corrs = []
            for k in range(len(sigmas)):
                S = sigmas[k]
                vols = np.sqrt(np.diag(S))
                inv_v = 1.0 / (vols + 1e-16)
                R = S * np.outer(inv_v, inv_v)
                corrs.append(R)
            return corrs
        
        # DYNAMIC CASE (T, K, D, D)
        else:
            T, K, D, _ = sigmas.shape
            corrs_dynamic = np.zeros_like(sigmas)
            for t in range(T):
                for k in range(K):
                    S = sigmas[t, k]
                    vols = np.sqrt(np.diag(S))
                    inv_v = 1.0 / (vols + 1e-16)
                    corrs_dynamic[t, k] = S * np.outer(inv_v, inv_v)
            return corrs_dynamic
        
    def compute_confidence_index(self, series_index: int = 0):
        """
        Computes the model's confidence index at each timestamp using Shannon Entropy.
        
        The index is normalized:
        - 1.0: Absolute certainty (one regime has 100% probability).
        - 0.0: Total uncertainty (all regimes have a probability of 1/K).

        Args:
            series_index (int): Index of the sequence to analyze.

        Returns:
            np.ndarray: Vector of confidence values in the range [0, 1].
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")
            
        # 1. Retrieve probabilities (T, K)
        probs = self.regime_probabilities(series_index)
        
        # 2. Compute Entropy H = -sum(p * log(p))
        # Add epsilon to prevent log(0)
        probs = np.maximum(probs, 1e-12)
        entropy = -np.sum(probs * np.log(probs), axis=1)
        
        # 3. Normalization: maximum entropy for K states is log(K)
        # Convert entropy to Confidence Index (0 to 1)
        max_entropy = np.log(self.K)
        confidence = 1 - (entropy / max_entropy)
        
        return confidence

    def get_params(self):
        """
        Retrieves the underlying model parameters (Initial dist, Transitions, Observations).

        Returns:
            tuple: Complete SSM parameter tuple.
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")
        return self.model.params # Returns a full ssm tuple

    def set_params(self, params):
        """
        Injects specific parameters into the model.

        Args:
            params (tuple): Parameters to be applied to the SSM model.
        """
        self.model.params = params
        self.is_fitted = True

    def get_transition_params(self):
        """
        Retrieves transition-specific parameters (log_Ps and Ws).

        If M = 0 (Classic HMM), returns log_Ps and a zero vector for Ws. 
        This ensures the intercept (log_Ps) remains the sole transition driver.

        Returns:
            tuple: (log_Ps, Ws) where log_Ps are the intercepts and Ws are covariate weights.
        """
        log_Ps = self.model.transitions.log_Ps.copy()
        if self.M > 0: # Exogenous covariates exist
            Ws = self.model.transitions.Ws.copy()
        else:
            # Return empty matrix (K x 0) to avoid breaking downstream concatenations
            Ws = np.zeros((self.K, self.M)) 
        return log_Ps, Ws

    def update_transition_params(self, log_Ps: np.ndarray, Ws: np.ndarray):
        """
        Updates the transition parameters (intercepts and weights).

        Weights (Ws) are only updated if exogenous covariates (M > 0) are present.

        Args:
            log_Ps (np.ndarray): Matrix (K x K) of base transition log-probabilities (intercepts).
            Ws (np.ndarray): Matrix (K x M) of weights associated with macro drivers.
        """
        self.model.transitions.log_Ps = log_Ps
        if self.M > 0: # Model has covariates, update weights
            self.model.transitions.Ws = Ws

    def compute_ll(self, Y: List[np.ndarray], X: List[np.ndarray]):
        """
        Computes the log-likelihood of the provided dataset.

        Args:
            Y (List[np.ndarray]): Observation sequences.
            X (List[np.ndarray]): Input/Covariate sequences.

        Returns:
            float: The calculated log-likelihood.
        """
        return self.model.log_likelihood(Y, inputs=X)

    def get_transition_matrices(self, X: np.ndarray = None):
        """
        Computes transition matrices based on model type:
        - If M == 0: Stationary model, X is ignored.
        - If M > 0: Macro-driven model (IOHMM), X is mandatory.

        Args:
            X (np.ndarray, optional): Exogenous inputs for time-varying transition calculation.

        Returns:
            np.ndarray: Matrix (K x K) or array of matrices (T x K x K).
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")

        if self.M == 0: # Stationary HMM
            P = self.model.transitions.transition_matrix
            return P

        else: # M > 0 case (IOHMM Macro)
            if X is None:
                raise ValueError(f"Model fitted with M={self.M} macro variables; X must be provided to compute transitions.")

            X_input = X
            T = X_input.shape[0]
            # Create a minimal dummy Y to satisfy ssm technical signature
            dummy_Y = np.zeros((T, self.D)) 

            return self.model.transitions.transition_matrices( # Follows ssm signature
                data=dummy_Y,
                input=X_input,
                mask=np.ones((T, self.D), dtype=bool),
                tag=None
            )
        
    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None, oos_start: int = 0, oos_end: int = None):
        """
        Computes predictive probabilities for both In-Sample and OOS periods, 
        returning only the OOS portion.

        Calculates at step t: P(z_t | Y_{1:t-1}, X_{1:t-1}).

        Args:
            observations (np.ndarray): Observations (log returns) for context + OOS.
            inputs (np.ndarray, optional): Covariates for context + OOS.
            oos_start (int): Start index for the Out-of-Sample period.
            oos_end (int, optional): End index for the Out-of-Sample period.

        Returns:
            np.ndarray: Predictive probabilities for the OOS period.
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")

        # all_preds[t] = P(z_t | Y_{1:t-1}, X_{1:t-1}) cf ssm documentation
        all_preds = self.model.filter(observations, input=inputs) 

        return all_preds[oos_start:oos_end]
import numpy as np
import ssm
from src.stock_bond_correlation.models.base_detector import BaseRegimeDetector


class HMMMacroRegimeDetector(BaseRegimeDetector):
    """
    Gaussian HMM (Classic Hidden Markov Model) based on the ssm library (Linderman Lab).

    The model is fitted on observation variables (e.g., standardized macro variables).
    Regime metrics (means, covariances) are computed on financial log returns,
    either in static mode (empirical by Viterbi state) or dynamic mode
    (EWMA weighted by regime probabilities).

    Parameters
    ----------
    n_states : int
        Number of latent regimes.
    n_iter : int
        Number of EM iterations for training.
    random_state : int
        Random seed for reproducibility of the final model.
    kappa : float
        Regime persistence strength ("sticky" transitions).
        0 = standard unbiased transitions, >0 = increasing persistence bias.
    n_restarts : int
        Number of K-Means initialized restarts.
        The selected model is the one maximizing the log-likelihood.
    """

    def __init__(self, n_states: int = 2, n_iter: int = 100, random_state: int = 42,
                 kappa: float = 5.0, n_restarts: int = 10):
        self.n_states = n_states
        self.n_iter = n_iter
        self.random_state = random_state
        self.seed = random_state            # alias kept for internal compatibility
        self.kappa = kappa
        self.n_restarts = n_restarts

        # Interface attributes expected by TransitionStatisticalAnalyzer / CoherenceChecker
        self.K = n_states                   # number of regimes (alias of n_states)
        self.D = None                       # observation dimension (set during fit)
        self.M = 0                          # number of covariate inputs (0 = unconditional HMM)

        self.model = None
        self.is_fitted = False              # simple boolean, set to True in fit()
        self._fitted_probs = None           # cache of smoothed probabilities (in-sample)
        self.viterbi_states = None          # most likely state sequence (Viterbi)
        self._observations_fitted = None    # training observations (e.g., macro variables)
        self._log_returns_fitted = None     # financial log returns aligned with fit
        self.fit_observations = None        # list-wrapped alias expected by TransitionStatisticalAnalyzer
        self.fit_input = None               # covariate inputs (None for unconditional HMM)

    # ------------------------------------------------------------------
    # Private method: instantiate SSM model
    # ------------------------------------------------------------------

    def _build_model(self, D: int):
        """
        Instantiates an ssm HMM according to the object's configuration.

        Parameters
        ----------
        D : int
            Observation dimension.
        """
        if self.kappa > 0:
            return ssm.HMM(
                self.n_states, D,
                observations="gaussian",
                transitions="sticky",
                transition_kwargs=dict(alpha=1.0, kappa=self.kappa)
            )
        return ssm.HMM(self.n_states, D, observations="gaussian")

    # ------------------------------------------------------------------
    # Private method: resolve log returns
    # ------------------------------------------------------------------

    def _resolve_log_returns(self, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Resolves the log returns to use: explicit argument or fitted cache.

        Parameters
        ----------
        log_returns : np.ndarray or None

        Returns
        -------
        np.ndarray (T, D)
        """
        if log_returns is not None:
            return np.array(log_returns, dtype=float)
        if self._log_returns_fitted is not None:
            return self._log_returns_fitted
        raise ValueError(
            "No log_returns available. "
            "Pass log_returns to fit() or directly to this method."
        )

    # ------------------------------------------------------------------
    # Private method: compute EWMA weights
    # ------------------------------------------------------------------

    def _ewma_weights(self, t: int, regime_probs: np.ndarray, halflife: int) -> np.ndarray:
        """
        Computes combined EWMA weights for regime k up to time t.

        Weights = exponential decay × probability of being in regime k.

        Parameters
        ----------
        t : int
            Current time step (exclusive).
        regime_probs : np.ndarray (t, K)
            Regime probabilities up to t.
        halflife : int
            Half-life in months for EWMA decay.

        Returns
        -------
        np.ndarray (t, K): normalized weights for each regime.
        """
        alpha = 1 - np.exp(-np.log(2) / halflife)
        decay = np.array([(1 - alpha) ** i for i in range(t)])[::-1]  # (t,)

        # Combined weights: decay × regime probability
        w = decay[:, None] * regime_probs[:t]  # (t, K)

        # Regime-wise normalization
        w_sum = w.sum(axis=0, keepdims=True)  # (1, K)
        w_sum = np.where(w_sum < 1e-8, 1.0, w_sum)
        return w / w_sum  # (t, K)

    # ------------------------------------------------------------------
    # fit
    # ------------------------------------------------------------------

    def fit(self, observations, inputs=None, log_returns=None):
        """
        Trains the HMM on observations using the EM algorithm with K-Means initialization.

        Performs n_restarts independent runs and retains the model maximizing
        the log-likelihood in order to avoid local optima.

        Parameters
        ----------
        observations : List[np.ndarray] or np.ndarray
            HMM observation variables (e.g., standardized macro variables).
            Accepts either a list of the form [obs_array] or an array (T, D).
            If a list is provided, observations[0] is used for fitting
            and observations[1] (if present) is stored as financial log returns.
        inputs : ignored
            HMMDetector is an unconditional transition model.
        log_returns : List[np.ndarray] or np.ndarray, optional
            Financial log returns aligned with observations.
            Stored for regime metric computation (means, covariances).
            Takes precedence over observations[1] if both are provided.
        """
        # Extraction if list format:
        # observations[0] → HMM fit data
        # observations[1] → financial log returns (optional, if log_returns not provided)
        if isinstance(observations, list):
            if log_returns is None and len(observations) > 1:
                log_returns = observations[1]
            observations = observations[0]
        observations = np.array(observations, dtype=float)
        if observations.ndim == 1:
            observations = observations[:, None]

        # Store financial log returns if provided
        if log_returns is not None:
            lr = log_returns[0] if isinstance(log_returns, list) else log_returns
            self._log_returns_fitted = np.array(lr, dtype=float)
        else:
            self._log_returns_fitted = None

        T, D = observations.shape

        # Multi-restart K-Means: retain the model with highest log-likelihood
        best_ll = -np.inf

        for seed in range(self.n_restarts):
            np.random.seed(seed)
            candidate = self._build_model(D)
            candidate.fit(observations, method="em", num_iters=self.n_iter,
                          init_method="kmeans", verbose=0)
            ll = candidate.log_probability(observations)
            if ll > best_ll:
                best_ll = ll
                self.model = candidate

        # Smoothed probabilities P(z_t | Y_{1:T}) via Forward-Backward
        self._fitted_probs = self.model.expected_states(observations)[0]

        # Optimal state sequence via the Viterbi algorithm
        self.viterbi_states = self.model.most_likely_states(observations)

        self._observations_fitted = observations
        self.fit_observations = [observations]  # list format expected by TransitionStatisticalAnalyzer
        self.fit_input = None                    # unconditional HMM: no inputs
        self.D = D          # observation dimension (available post-fit)
        self.K = self.n_states  # redundant but ensures consistency if n_states changes
        self.is_fitted = True
        return self

    # ------------------------------------------------------------------
    # regime_probabilities
    # ------------------------------------------------------------------

    def regime_probabilities(self, series_index: int = 0) -> np.ndarray:
        """
        Returns in-sample smoothed probabilities P(z_t | Y_{1:T}) — shape (T, K).

        Parameters
        ----------
        series_index : ignored (kept for compatibility with abstract contract).
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")
        return self._fitted_probs

    # ------------------------------------------------------------------
    # regime_covariances — static or dynamic EWMA
    # ------------------------------------------------------------------

    def regime_covariances(self, log_returns: np.ndarray = None,
                           use_ewma: bool = True,
                           halflife: int = 60):
        """
        Returns regime covariance matrices of log returns.

        Static mode (use_ewma=False):
            Fixed empirical covariances by Viterbi state — shape (K, D, D).
            Static case in the sense of base.py: fixed parameters estimated during fit.

        Dynamic EWMA mode (use_ewma=True):
            EWMA covariances weighted by regime probabilities — shape (T, K, D, D).
            Dynamic case in the sense of base.py: time sequence of matrices.
            Each matrix C[t, k] is estimated on the history up to t,
            exponentially weighted and weighted by P(z_t = k).

        Parameters
        ----------
        log_returns : np.ndarray, optional (T, D)
            Financial log returns. If None, uses those stored during fit.
        use_ewma : bool
            True → dynamic EWMA mode, False → empirical static mode.
        halflife : int
            Half-life in months for EWMA decay (used if use_ewma=True).
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")

        lr = self._resolve_log_returns(log_returns)
        T, D = lr.shape
        K = self.n_states

        if not use_ewma:
            # Static mode: empirical covariance by Viterbi state
            covs = []
            for k in range(K):
                mask = (self.viterbi_states == k)
                subset = lr[mask]
                covs.append(np.cov(subset.T, ddof=1) if subset.shape[0] >= 2 else np.eye(D))
            return covs

        # Dynamic EWMA mode — shape (T, K, D, D)
        covs_dynamic = np.zeros((T, K, D, D))

        for t in range(1, T):
            w = self._ewma_weights(t, self._fitted_probs, halflife)  # (t, K)

            for k in range(K):
                w_k = w[:, k]  # (t,)

                # Weighted mean
                mu_k = (lr[:t] * w_k[:, None]).sum(axis=0)

                # Weighted covariance
                diff = lr[:t] - mu_k
                covs_dynamic[t, k] = (diff * w_k[:, None]).T @ diff

        return covs_dynamic

    # ------------------------------------------------------------------
    # regime_means — static or dynamic EWMA
    # ------------------------------------------------------------------

    def regime_means(self, log_returns: np.ndarray = None,
                     use_ewma: bool = True,
                     halflife: int = 24):
        """
        Returns regime means of log returns.

        Static mode (use_ewma=False):
            Fixed empirical means by Viterbi state — shape (K, D).

        Dynamic EWMA mode (use_ewma=True):
            EWMA means weighted by regime probabilities — shape (T, K, D).

        Parameters
        ----------
        log_returns : np.ndarray, optional (T, D)
            Financial log returns. If None, uses those stored during fit.
        use_ewma : bool
            True → dynamic EWMA mode, False → empirical static mode.
        halflife : int
            Half-life in months for EWMA decay (used if use_ewma=True).
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")

        lr = self._resolve_log_returns(log_returns)
        T, D = lr.shape
        K = self.n_states

        if not use_ewma:
            # Static mode: empirical mean by Viterbi state
            means = []
            for k in range(K):
                mask = (self.viterbi_states == k)
                subset = lr[mask]
                means.append(np.mean(subset, axis=0) if subset.shape[0] > 0 else np.zeros(D))
            return means

        # Dynamic EWMA mode — shape (T, K, D)
        means_dynamic = np.zeros((T, K, D))

        for t in range(1, T):
            w = self._ewma_weights(t, self._fitted_probs, halflife)  # (t, K)

            for k in range(K):
                w_k = w[:, k]  # (t,)
                means_dynamic[t, k] = (lr[:t] * w_k[:, None]).sum(axis=0)

        return means_dynamic

    # ------------------------------------------------------------------
    # conditional_covariance — delegation to BaseRegimeDetector
    # ------------------------------------------------------------------

    def conditional_covariance(self, probs: np.ndarray,
                                log_returns: np.ndarray = None) -> np.ndarray:
        """
        Computes total conditional covariance via the law of total covariance.

        Delegates to BaseRegimeDetector.conditional_covariance while providing
        regime metrics computed on log returns.

        Parameters
        ----------
        probs : np.ndarray (T, K)
            Regime probabilities at each date.
        log_returns : np.ndarray, optional
            Financial log returns (T, D). If None, uses those stored during fit.
        """
        lr = self._resolve_log_returns(log_returns)
        return super().conditional_covariance(probs=probs, log_returns=lr)

    # ------------------------------------------------------------------
    # predict_probabilities
    # ------------------------------------------------------------------

    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None,
                              oos_start: int = 0, oos_end: int = None) -> np.ndarray:
        """
        Computes causal predictive probabilities P(z_t | Y_{1:t-1}) on OOS data.

        Uses Forward filtering (causal) — no look-ahead bias.

        Parameters
        ----------
        observations : np.ndarray
            Full in-sample + OOS observations (T_total, D).
        inputs : ignored
        oos_start : int
            OOS start index in observations.
        oos_end : int, optional
            OOS end index. If None, goes until the end.

        Returns
        -------
        np.ndarray (T_oos, K): predictive probabilities on OOS only.
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")

        observations = np.array(observations, dtype=float)
        if observations.ndim == 1:
            observations = observations[:, None]

        if oos_end is None:
            oos_end = len(observations)

        # Causal filtering on full sequence (in-sample + OOS for context)
        _, filtered = self.model.filter(observations)

        # Return only filtered probabilities on OOS
        return filtered[oos_start:oos_end]

    # ------------------------------------------------------------------
    # get_transition_matrix
    # ------------------------------------------------------------------

    def get_transition_matrix(self) -> np.ndarray:
        """
        Returns the stochastic transition matrix A (K, K),
        where A[i, j] = P(z_{t+1} = j | z_t = i).
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")
        return np.exp(self.model.transitions.log_Ps)

    def get_transition_matrices(self, inputs=None) -> np.ndarray:
        """
        Returns the sequence of transition matrices — shape (T, K, K).

        For a standard (unconditional) HMM, the matrix is constant:
        the same matrix (K, K) is repeated T times. If inputs are provided,
        T is inferred from their length, otherwise from fitted observations.

        Parameters
        ----------
        inputs : np.ndarray (T, M) or None — ignored for an unconditional HMM.

        Returns
        -------
        np.ndarray (T, K, K)
        """
        A = self.get_transition_matrix()  # (K, K)
        if inputs is not None:
            T = np.array(inputs).shape[0]
        elif self._observations_fitted is not None:
            T = self._observations_fitted.shape[0]
        else:
            raise ValueError("Unable to infer T: pass inputs or call fit() first.")
        return np.tile(A, (T, 1, 1))  # (T, K, K)

    # ------------------------------------------------------------------
    # get_transition_params / set_transition_params
    # Required by TransitionStatisticalAnalyzer (save → perturb → restore)
    # ------------------------------------------------------------------

    def get_transition_params(self):
        """
        Returns raw transition parameters for backup.

        For a standard (unconditional) HMM, only transition log-probabilities
        log_Ps are free parameters. Ws is returned as
        an empty array of shape (K, 0) — without covariates — to remain compatible
        with TransitionStatisticalAnalyzer, which always assumes an indexable array.

        Returns
        -------
        log_Ps : np.ndarray (K, K) — copy of transition log-probabilities.
        Ws     : np.ndarray (K, 0) — empty array (no covariates).
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")
        log_Ps = self.model.transitions.log_Ps.copy()
        Ws = np.zeros((self.K, 0))   # (K, M=0): compatible with all array ops
        return log_Ps, Ws

    def set_transition_params(self, log_Ps: np.ndarray, Ws=None):
        """
        Restores transition parameters (after numerical perturbation).

        Parameters
        ----------
        log_Ps : np.ndarray (K, K) — transition log-probabilities to restore.
        Ws     : ignored (no covariates in a standard HMM).
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")
        self.model.transitions.log_Ps = log_Ps.copy()

    def update_transition_params(self, log_Ps: np.ndarray, Ws=None):
        """
        Alias of set_transition_params — used by TransitionStatisticalAnalyzer
        during numerical perturbations for Hessian computation.

        Parameters
        ----------
        log_Ps : np.ndarray (K, K) — updated transition log-probabilities.
        Ws     : ignored (no covariates in a standard HMM).
        """
        self.set_transition_params(log_Ps, Ws)

    def compute_ll(self, observations_list, inputs_list=None) -> float:
        """
        Computes model log-likelihood on the provided observations.

        Used by TransitionStatisticalAnalyzer to evaluate the LL after
        perturbation of transition parameters (numerical Hessian computation).

        Parameters
        ----------
        observations_list : list[np.ndarray]
            List containing an observation array (T, D).
        inputs_list : ignored — unconditional HMM.

        Returns
        -------
        float : model log-likelihood.
        """
        if self.model is None:
            raise ValueError("The model has not been trained.")
        obs = np.array(observations_list[0], dtype=float)
        if obs.ndim == 1:
            obs = obs[:, None]
        return self.model.log_probability(obs)

    # ------------------------------------------------------------------
    # compute_confidence_index — required by HMMCoherenceChecker
    # ------------------------------------------------------------------

    def compute_confidence_index(self, series_index: int = 0) -> np.ndarray:
        """
        Computes the confidence index at each time step as
        (1 - normalized entropy).

        A value close to 1 indicates that the model is highly certain
        about the current regime (one probability dominates).
        A value close to 0 indicates strong ambiguity.

        Parameters
        ----------
        series_index : int
            Ignored — kept for compatibility with BaseRegimeDetector contract.

        Returns
        -------
        np.ndarray (T,) : confidence index ∈ [0, 1] for each date.
        """
        if self._fitted_probs is None:
            raise ValueError("The model has not been trained.")

        probs = self._fitted_probs  # (T, K)
        K = probs.shape[1]

        # Normalized Shannon entropy: H ∈ [0, 1]
        # Clip to avoid log(0)
        log_probs = np.log(np.clip(probs, 1e-12, 1.0))
        entropy = -np.sum(probs * log_probs, axis=1)          # (T,)
        max_entropy = np.log(K)                                # max entropy = log(K) (uniform)
        normalized_entropy = entropy / max_entropy             # (T,) ∈ [0, 1]

        return 1.0 - normalized_entropy                        # (T,) ∈ [0, 1]

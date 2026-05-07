import numpy as np
import ssm 
from RegimeDetector.base_detector import BaseRegimeDetector
from typing import List

class HMMMarketRegimeDetector(BaseRegimeDetector):
    """
    Wrapper for the 'ssm' library dedicated to detecting market regimes.

    Generic Logic:
    Fits an HMM on historical data up to time T (Y_{1:t}, X_{1:t}) and then 
    predicts the hidden state (regime) for T+1.

    Note:
        This model remains functional even without exogenous covariates/inputs X 
        (i.e., X is None and the transition matrix is constant/sticky). 
        The 'ssm' library is flexible enough to ignore null inputs.
    """
    def __init__(self, n_regimes: int = 2, n_dim: int = 2, observations_type: str = "gaussian", transitions: str = "sticky",
                 transition_kwargs: dict = None, random_state: int = 42, n_input : int = 0):
        """
        Initializes the Market Regime Detector.

        Args:
            n_regimes (int): Desired number of regimes. Defaults to 2.
            n_dim (int): Number of assets whose log returns contribute to market regimes. Defaults to 2.
            observations_type (str): Distribution choice for log returns in each regime. Defaults to "gaussian".
            transitions (str): Transition matrix modeling. Use "sticky" for a persistent matrix, otherwise "standard".
            transition_kwargs (dict): Hyperparameters for transitions (e.g., kappa and alpha for sticky HMM).
            random_state (int): Seed for the EM algorithm initialization and execution. Defaults to 42.
            n_input (int): Number of exogenous inputs impacting transitions. Defaults to 0 for this static model.
        """
        self.K = n_regimes
        self.D = n_dim
        self.M = n_input
        self.observations_type = observations_type
        self.transitions = transitions
        # Default sticky parameters: kappa=10, alpha=2
        self.transition_kwargs = transition_kwargs or {"kappa": 10, "alpha": 2}
        self.random_state = random_state

        np.random.seed(self.random_state)
        # Initialize the underlying ssm model
        self.model = ssm.HMM(self.K, self.D, self.M, observations=self.observations_type, transitions=self.transitions,
                             transition_kwargs=self.transition_kwargs) 

        self.is_fitted = False
        self.fit_observations = None
        self.fit_input = None
        self.viterbi_states = None # Represents the hidden states on the training set (in-sample)

    def fit(self, observations: List[np.ndarray] = None,  inputs: List[np.ndarray] = None, log_returns : List[np.ndarray] = None,
            series_index: int = 0, num_iters: int = 200, initialize: bool = True, init_method : str = "kmeans"):
        """
        Fits the initialized model and estimates the hidden states for each timestamp.

        Args:
            observations (List[np.ndarray]): List of log return sequences from 1...T.
            inputs (List[np.ndarray], optional): List of covariates from 1...T (included for IOHMM inheritance, ignored here).
            log_returns (List[np.ndarray], optional): Ignored in this model (included for API consistency).
            series_index (int): Index of the specific sequence in the observations list to fit. Defaults to 0.
            num_iters (int): Number of Expectation-Maximization (EM) iterations. Defaults to 200.
            initialize (bool): Whether to initialize parameters before fitting. Defaults to True.
            init_method (str): Method used for initialization (e.g., "kmeans").
        """
        np.random.seed(self.random_state)
        self.model.fit(observations, method="em", num_iters=num_iters, init_method=init_method, verbose=0, initialize=initialize)
        self.is_fitted = True
        self.fit_observations = observations
        self.fit_input = inputs
        # Compute the most likely states for each timestamp using the Viterbi algorithm (In-Sample)
        self.viterbi_states = self.model.most_likely_states(observations[series_index], input=inputs[series_index] if inputs is not None else None)
        return self

    def regime_probabilities(self, series_index : int = 0):
        """
        Calculates the smoothed probabilities for each regime at each date using the 'expected_states' method.

        Args:
            series_index (int): Index of the sequence to analyze.

        Returns:
            np.ndarray: Matrix of shape (T, K) containing probabilities.
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")
        data = self.fit_observations[series_index] # array
        input = self.fit_input[series_index] if self.fit_input is not None else None
        # expected_states returns a tuple, we take the first element (the states)
        return self.model.expected_states(data=data, input=input)[0] 

    def regime_covariances(self, log_returns : np.ndarray = None):
        """
        Returns the list of covariance matrices for each identified regime.

        Note:
            In this static model, covariance matrices are constant within each regime; 
            the 'log_returns' argument is ignored but kept for general API compatibility.

        Returns:
            List[np.ndarray]: List of K covariance matrices of shape (D, D).
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")

        cov_array = self.model.observations.Sigmas
        return [cov_array[k] for k in range(self.K)]

    def regime_means(self, log_returns: np.ndarray = None):
        """
        Returns the mean vectors of log returns for each identified regime.

        Note:
            In this static model, mean returns are constant within each regime; 
            the 'log_returns' argument is ignored but kept for general API compatibility.

        Returns:
            List[np.ndarray]: List of K mean vectors of shape (D,).
        """
        if not self.is_fitted:
            raise ValueError("Model not fitted.")
        # Convert to list to respect the abstract contract
        return [self.model.observations.mus[k] for k in range(self.K)]
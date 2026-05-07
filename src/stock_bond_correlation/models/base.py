from abc import ABC, abstractmethod
import numpy as np
from typing import List, Union


class RegimeDetector(ABC):
    """
    Abstract base class defining the mandatory interface for all regime detectors.

    Role:
    - Ensure a consistent API across different HMM implementations.
    - Provide methods for model fitting, probability estimation, and risk parameter extraction.
    - Standardize the handling of both static and dynamic (EWMA-based) regime parameters.

    Purpose:
    Enforce essential method implementation to prevent API drift between various 
    detector versions (Market-only, Macro-driven, TVTP, etc.).
    """

    def __init__(self, n_regimes: int, n_dim: int, n_input: int = 0, random_state: int = 42):
        """
        Initializes the base regime detector parameters.

        Args:
            n_regimes (int): Number of hidden states (regimes) to identify.
            n_dim (int): Number of assets or dimensions in the observation space.
            n_input (int, optional): Number of exogenous covariates/inputs impacting transitions. Defaults to 0.
            random_state (int, optional): Random seed for reproducibility. Defaults to 42.

        Note:
            Subclasses must initialize their specific HMM library model as self.model.
        """
        self.K = n_regimes
        self.D = n_dim
        self.M = n_input
        self.random_state = random_state
        
        # Status flags and data storage
        self.is_fitted = False
        self.viterbi_states = None
        self.fit_observations = None
        self.fit_input = None

    @abstractmethod
    def fit(self, observations: List[np.ndarray], inputs: List[np.ndarray] = None, log_returns: List[np.ndarray] = None,
            initialize: bool = True, init_method: str = "kmeans", num_iters: int = 200):
        """
        Fits the HMM model to the provided historical data.

        Args:
            observations (List[np.ndarray]): List of observation arrays (typically a list of size 1 for time-series).
            inputs (List[np.ndarray], optional): List of covariate arrays affecting transition probabilities.
            log_returns (List[np.ndarray], optional): Log returns of the assets (used if different from observations).
            initialize (bool): Whether to perform parameter initialization before fitting. Defaults to True.
            init_method (str): Strategy for parameter initialization (e.g., "kmeans", "random"). Defaults to "kmeans".
            num_iters (int): Maximum number of iterations for the EM algorithm. Defaults to 200.
        """
        pass

    @abstractmethod
    def regime_probabilities(self, series_index: int = 0) -> np.ndarray:
        """
        Returns the in-sample filtered or smoothed probabilities for each regime.

        Args:
            series_index (int): Index of the sequence within the observations list to analyze.

        Returns:
            np.ndarray: Array of shape (T, K) containing probabilities of being in each regime at time t.
        """
        pass

    @abstractmethod
    def regime_covariances(self, log_returns: np.ndarray = None) -> Union[List[np.ndarray], np.ndarray]:
        """
        Extracts asset covariance matrices for each identified regime.

        Args:
            log_returns (np.ndarray, optional): Required for dynamic models to compute time-varying covariance.

        Returns:
            Union[List[np.ndarray], np.ndarray]: 
                - Static models: List of K matrices of shape (D, D).
                - Dynamic models: Array of shape (T, K, D, D) representing time-varying parameters.
        """
        pass

    @abstractmethod
    def regime_means(self, log_returns: np.ndarray = None) -> Union[List[np.ndarray], np.ndarray]:
        """
        Extracts asset mean returns for each identified regime.

        Args:
            log_returns (np.ndarray, optional): Required for dynamic models to compute time-varying means.

        Returns:
            Union[List[np.ndarray], np.ndarray]:
                - Static models: List of K vectors of shape (D,).
                - Dynamic models: Array of shape (T, K, D) representing time-varying parameters.
        """
        pass
    
    @abstractmethod
    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None, 
                             oos_start: int = 0, oos_end: int = None) -> np.ndarray:
        """
        Computes one-step-ahead predictive probabilities for the Out-of-Sample (OOS) period.

        Calculates P(z_t | Y_{1:t-1}, X_{1:t-1}) using historical context for initialization.

        Args:
            observations (np.ndarray): Full observation array (In-sample + OOS for context).
            inputs (np.ndarray, optional): Full covariate array (In-sample + OOS).
            oos_start (int): Starting index of the Out-of-Sample period.
            oos_end (int, optional): Ending index of the Out-of-Sample period.

        Returns:
            np.ndarray: Predictive probabilities for the OOS period only.
        """
        pass

    @abstractmethod
    def conditional_covariance(self, probs: np.ndarray, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Computes the expected covariance matrices by weighting regime parameters with probabilities.

        Args:
            probs (np.ndarray): Probability array (typically predictive) used for weighting.
            log_returns (np.ndarray, optional): Required for dynamic models.

        Returns:
            np.ndarray: Array of shape (T_oos, D, D) representing the conditional covariance at each step.
        """
        pass

    @abstractmethod 
    def get_params(self):
        """
        Retrieves the model parameters (Initial distribution, Transition matrix, Observation parameters).

        Returns:
            dict or object: Current model state parameters.
        """
        pass

    @abstractmethod 
    def set_params(self, params):
        """
        Injects a specific set of parameters into the model.

        Args:
            params (dict or object): Parameters to be applied to the model.
        """
        pass
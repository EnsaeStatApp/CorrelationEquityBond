import numpy as np
import ssm
from RegimeDetector.base_detector import BaseRegimeDetector

class HMMDetector(BaseRegimeDetector):

    def __init__(self, n_states: int = 2, n_iter: int = 100, random_state: int = 42):
        self.n_states = n_states
        self.n_iter = n_iter
        self.seed = random_state
        self.model = None
        self._fitted_probs = None
        self.viterbi_states = None

    def fit(self, Y, X=None):
        if isinstance(Y, list):
            Y = Y[0]
        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]
        T, D = Y.shape
        np.random.seed(self.seed)
        self.model = ssm.HMM(self.n_states, D, observations="gaussian")
        self.model.fit(Y, method="em", num_iters=self.n_iter, verbose=0)
        self._fitted_probs = self.model.expected_states(Y)[0]
        self.viterbi_states = self.model.most_likely_states(Y)
        self._Y_fitted = Y
        return self

    def regime_probabilities(self, series_index=0, Y=None, X=None):
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        if Y is None:
            return self._fitted_probs
        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]
        return self.model.expected_states(Y)[0]

    def regime_covariances(self):
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        return self.model.observations.Sigmas

    def get_transition_matrix(self):
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        return np.exp(self.model.transitions.log_Ps)

    def predict_probabilities(self, Y, X=None, horizon=1):
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]
        _, filtered = self.model.filter(Y)
        current_probs = filtered[-1]
        A = self.get_transition_matrix()
        for _ in range(horizon):
            current_probs = current_probs @ A
        return current_probs

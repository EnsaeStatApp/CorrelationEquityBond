import numpy as np
import ssm
from RegimeDetector.base_detector import BaseRegimeDetector


class HMMDetector(BaseRegimeDetector):
    """
    HMM Gaussien standard (Classic Hidden Markov Model) basé sur la librairie ssm (Linderman Lab).

    Paramètres
    ----------
    n_states : int
        Nombre de régimes latents.
    n_iter : int
        Nombre d'itérations EM pour l'apprentissage.
    random_state : int
        Graine aléatoire pour la reproductibilité.
    """

    def __init__(self, n_states: int = 2, n_iter: int = 100, random_state: int = 42):
        self.n_states = n_states
        self.n_iter = n_iter
        self.seed = random_state
        self.model = None
        self._fitted_probs = None   # Cache des probabilités lissées (in-sample)
        self.viterbi_states = None  # Séquence d'états la plus probable (Viterbi)
        self._Y_fitted = None       # Données d'entraînement conservées pour référence

    def fit(self, Y, X=None):
        """
        Entraîne le HMM sur les données Y via l'algorithme EM.

        Accepte Y sous forme de liste (contrat RegimeDetector) ou de tableau (T, D).
        X est ignoré : HMMDetector est un modèle non-conditionnel.
        """
        # Extraction du tableau si Y est passé sous forme de liste
        if isinstance(Y, list):
            Y = Y[0]

        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]  # (T,) → (T, 1)

        T, D = Y.shape

        np.random.seed(self.seed)
        self.model = ssm.HMM(self.n_states, D, observations="gaussian")

        # Apprentissage par Expectation-Maximisation
        self.model.fit(Y, method="em", num_iters=self.n_iter, verbose=0)

        # Probabilités lissées P(z_t | Y_{1:T}) via Forward-Backward
        self._fitted_probs = self.model.expected_states(Y)[0]

        # Séquence optimale d'états via l'algorithme de Viterbi
        self.viterbi_states = self.model.most_likely_states(Y)

        self._Y_fitted = Y
        return self

    def regime_probabilities(self, series_index=0, Y=None, X=None):
        """
        Retourne les probabilités lissées P(z_t | Y_{1:T}).

        Si Y est None, retourne les probabilités in-sample calculées lors du fit.
        Sinon, effectue l'inférence Forward-Backward sur les nouvelles données.
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        if Y is None:
            return self._fitted_probs

        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]

        return self.model.expected_states(Y)[0]

    def regime_covariances(self):
        """
        Retourne les matrices de covariance gaussiennes par régime — shape (K, D, D).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        return self.model.observations.Sigmas

    def get_transition_matrix(self):
        """
        Retourne la matrice de transition stochastique A (K, K),
        où A[i, j] = P(z_{t+1} = j | z_t = i).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        return np.exp(self.model.transitions.log_Ps)

    def predict_probabilities(self, Y, X=None, horizon=1):
        """
        Prédit le vecteur de probabilités d'état à T + horizon.

        Calcule P(z_T | Y_{1:T}) par filtrage, puis propage via A^horizon.

        Paramètres
        ----------
        Y : np.ndarray
            Série observée jusqu'à l'instant T.
        horizon : int
            Nombre de pas dans le futur.

        Retourne
        --------
        np.ndarray de shape (K,) : P(z_{T+horizon} | Y_{1:T})
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]

        # Probabilité filtrée au dernier instant observé
        _, filtered = self.model.filter(Y)
        current_probs = filtered[-1]

        # Propagation sur `horizon` pas via la matrice de transition
        A = self.get_transition_matrix()
        for _ in range(horizon):
            current_probs = current_probs @ A

        return current_probs

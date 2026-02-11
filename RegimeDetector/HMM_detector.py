import numpy as np
import ssm
#from .base_detector import BaseRegimeDetector
#Dans un environnement autre que Jupiter ou Google Collab, enlever le hashtag ci-dessus

class HMMDetector(BaseRegimeDetector):
    """
    Implémentation d'un HMM Gaussien standard (Classic Hidden Markov Model)
    basé sur la librairie 'ssm' (Linderman Lab).
    """

    def __init__(self, n_states: int = 2, n_iter: int = 100, random_state: int = 42):
        self.n_states = n_states
        self.n_iter = n_iter
        self.seed = random_state
        self.model = None
        self._fitted_probs = None # Cache pour les probas sur le train set

    def fit(self, Y: np.ndarray, X: np.ndarray = None):
        """
        Adapte un HMM Gaussien sur les données Y (T x D).
        X est ignoré pour un HMM standard (ce n'est pas un InputHMM).
        """
        # Assurer le format numpy float (T, D)
        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]

        T, D = Y.shape

        # Initialisation du modèle HMM via ssm
        np.random.seed(self.seed)
        # observations="gaussian" crée un GaussianHMM standard
        self.model = ssm.HMM(self.n_states, D, observations="gaussian")

        # Apprentissage (EM)
        self.model.fit(Y, method="em", num_iters=self.n_iter, verbose=0)

        # Calcul et stockage des probabilités lissées (smoothed probabilities) sur le train
        # expected_states retourne (Ez, Ezz, Ezzz), où Ez est (T, K)
        expectations = self.model.expected_states(Y)
        self._fitted_probs = expectations[0]

        return self

    def regime_probabilities(self, series_index: int = 0, Y: np.ndarray = None, X: np.ndarray = None):
        """
        Renvoie les probabilités d'état lissées P(z_t | Y_{1:T}).
        Si Y est None, renvoie celles calculées lors du fit.
        Sinon, calcule sur les nouvelles données Y.
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné (fit).")

        if Y is None:
            return self._fitted_probs

        # Inférence sur de nouvelles données
        Y = np.array(Y, dtype=float)
        if Y.ndim == 1:
            Y = Y[:, None]

        # On utilise expected_states (Forward-Backward) pour avoir les probas lissées
        expectations = self.model.expected_states(Y)
        return expectations[0]

    def regime_covariances(self):
        """
        Renvoie les matrices de covariance des observations pour chaque régime.
        Format : (K, D, D)
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")

        # Dans ssm, accès direct via model.observations.Sigmas
        return self.model.observations.Sigmas

    def get_transition_matrix(self):
        """
        Helper spécifique à SSM pour récupérer la matrice de transition (K x K).
        """
        if self.model is None:
            raise ValueError("Le modèle n'a pas été entraîné.")
        return np.exp(self.model.transitions.log_Ps)

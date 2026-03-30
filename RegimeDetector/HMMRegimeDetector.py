import numpy as np
import ssm as ssm 
from RegimeDetector.base_detector import BaseRegimeDetector

class HMMRegimeDetector(BaseRegimeDetector):
    """
    Wrappeur de la librairie ssm pour détecter les régimes de marché
    
    NB : Dans ce modèle, même si on ne rentre pas de covariables/inputs X (i.e. X est None et la matrice de transition est constante/sticky), 
    ce code fonctionne car ssm est suffisamment flexible pour l'ignorer
    """
    def __init__(self, n_regimes: int = 2, n_dim: int = 2, observations: str = "gaussian", transitions: str = "sticky", 
                 transition_kwargs: dict = None, random_state: int = 42, n_input : int = 0):
        """
        Parametres: 
        - n_regimes : nombre de régimes souhaités
        - n_dim : nombre d'actifs dont les log returns vont contribués aux régimes de marché
        - observations : choix de la distribution des log returns dans chaque régime (pour l'instant "gaussian")
        - transitions : choix de la modélisation de la matrice de transition ("sticky" pour matrice de transition collante, sinon "standard)
        - transition_kwargs : choix des hyperparamètres pour les transitions (kappa et alpha dans ce modèle)
        - random_state : le seed sur lequel va se faire l'algo EM (important)
        - n_input : nombre de covariables/inputs qui impactent les transitions (défaut 0 dans ce modèle statique)

        """
        self.K = n_regimes
        self.D = n_dim
        self.M = n_input
        self.observations = observations
        self.transitions = transitions
        self.transition_kwargs = transition_kwargs or {"kappa": 10, "alpha": 2}
        self.random_state = random_state
        
        np.random.seed(self.random_state)
        self.model = ssm.HMM(self.K, self.D, self.M, observations=self.observations, transitions=self.transitions, 
                             transition_kwargs=self.transition_kwargs) # modèle initialisé
        
        self.is_fitted = False
        self.fit_data = None
        self.viterbi_states = None # représente les régimes sur le jeu de train (in-sample)

    def fit(self, Y: list = None,  X: list = None, series_index: int = 0, num_iters: int = 200):
        """
        Fit le modèle initialisé et initialise les états pour chaque date 

        Paramètres:
        - Y : liste de séries de log returns (on ne s'intéresse en fait qu'à Y[series_index])
        - X : liste de covariables (présent par soucis de généralité mais ignoré dans ce modèle)
        - num_iters : nombres d'itérations de l'algo EM
        """
        np.random.seed(self.random_state)
        self.model.fit(Y, method="em", num_iters=num_iters, init_method="kmeans", verbose=0)
        self.is_fitted = True
        self.fit_data = Y
        self.fit_input = X
        self.viterbi_states = self.model.most_likely_states(Y[series_index], input=X[series_index] if X is not None else None) # les états les plus probables pour chaque date (NB : on utilise viterbi de ssm car on est sur l'in-sample)
        return self

    def regime_probabilities(self, series_index : int = 0):
        """
        Calcule les probabilités d'être dans chaque régime à chaque date via la méthode expected_states de ssm

        Paramètres:
        - Y : liste de séries de log returns (on ne s'intéresse en fait qu'à Y[series_index])
        - X : liste de covariables (présent par soucis de généralité mais ignoré dans ce modèle)
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
        data = self.fit_data[series_index] # array
        input = self.fit_input[series_index] if self.fit_input is not None else None
        return self.model.expected_states(data=data, input=input)[0] # [0] car renvoie un tuple

    def regime_covariances(self):
        """
        Retourne la liste des matrices de covariance pour chaque régime
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
    
        cov_array = self.model.observations.Sigmas
        return [cov_array[k] for k in range(self.K)]
    
    def predict_probabilities(self, Y: np.ndarray, X: np.ndarray = None):
        """
        Calcule les probabilités prédictives P(z_t | Y_{1:t-1})

        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
        
        predictive_probs = self.model.filter(data=Y, input=X) # predictive_probs[t] = P(z_t | Y_{1:t-1})
        return predictive_probs
    
    def regime_means(self):
        """
        Renvoie les moyennes des log returns dans chaque régime
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
        return self.model.observations.mus
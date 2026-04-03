import numpy as np
import ssm 
from RegimeDetector.base_detector import BaseRegimeDetector
from typing import List

class HMMMarketRegimeDetector(BaseRegimeDetector):
    """
    Wrappeur de la librairie ssm pour détecter les régimes de marché

    Fonctionnement générique : fit un HMM sur des données jusqu'à T (Y_{1:t}, X_{1:t}) puis prédit les régimes en T+1

    NB : Dans ce modèle, même si on ne rentre pas de covariables/inputs X (i.e. X est None et la matrice de transition est constante/sticky),
    ce code fonctionne car ssm est suffisamment flexible pour l'ignorer
    """
    def __init__(self, n_regimes: int = 2, n_dim: int = 2, observations_type: str = "gaussian", transitions: str = "sticky",
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
        self.observations_type = observations_type
        self.transitions = transitions
        self.transition_kwargs = transition_kwargs or {"kappa": 10, "alpha": 2}
        self.random_state = random_state

        np.random.seed(self.random_state)
        self.model = ssm.HMM(self.K, self.D, self.M, observations=self.observations_type, transitions=self.transitions,
                             transition_kwargs=self.transition_kwargs) # modèle initialisé

        self.is_fitted = False
        self.fit_observations = None
        self.fit_input = None
        self.viterbi_states = None # représente les régimes sur le jeu de train (in-sample)

    def fit(self, observations: List[np.ndarray] = None,  inputs: List[np.ndarray] = None, log_returns : List[np.ndarray] = None,
            series_index: int = 0, num_iters: int = 200):
        """
        Fit le modèle initialisé et initialise les états pour chaque date

        Paramètres:
        - observations : liste de séries de log returns de 1, ..., T (on ne s'intéresse en fait qu'à Y[series_index], mais ssm fit sur une liste)
        - inputs : liste de covariables de 1, ..., T (présent pour héritage IOHMM mais ignoré dans ce modèle)
        - log_returns : ignoré dans ce modèle car déjà dans observations (présent par soucis de généralité)
        - series_index : index de la liste sur lequel on fit le hmm (par défaut toujours 0)
        - num_iters : nombres d'itérations de l'algo EM
        """
        np.random.seed(self.random_state)
        self.model.fit(observations, method="em", num_iters=num_iters, init_method="kmeans", verbose=0)
        self.is_fitted = True
        self.fit_observations = observations
        self.fit_input = inputs
        self.viterbi_states = self.model.most_likely_states(observations[series_index], input=inputs[series_index] if inputs is not None else None) # les états les plus probables pour chaque date (NB : on utilise viterbi de ssm car on est sur l'in-sample)
        return self

    def regime_probabilities(self, series_index : int = 0):
        """
        Calcule les probabilités d'être dans chaque régime à chaque date via la méthode expected_states de ssm sur le in-sample

        Paramètres:
        - series_index : l'index de la série à étudier
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
        data = self.fit_observations[series_index] # array
        input = self.fit_input[series_index] if self.fit_input is not None else None
        return self.model.expected_states(data=data, input=input)[0] # [0] car renvoie un tuple

    def regime_covariances(self, log_returns : np.ndarray = None):
        """
        Retourne la liste des matrices de covariance pour chaque régime

        Paramètre :
        - log_returns : ignoré dans ce modèle les matrices de cov sont statiques dans chaque régime (mais présent par soucis de généralité)
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")

        cov_array = self.model.observations.Sigmas
        return [cov_array[k] for k in range(self.K)]

    def regime_means(self, log_returns: np.ndarray = None):
        """
        Renvoie les moyennes des log returns dans chaque régime

        Paramètre :
        - log_returns : ignoré dans ce modèle les moyennes sont stastiques dans chaque régime (mais présent par soucis de généralité)
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")
        return [self.model.observations.mus[k] for k in range(self.K)] # convertit en liste pour le contract abstrait


    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None, oos_start : int = 0, oos_end : int = None):
        """
        Calcule les probabilités prédictives sur tout l'in-sample + l'OOS (pour le contexte) mais ne renvoie les probabilités que
        sur l'OOS

        A la ligne t, P(z_t | Y_{1:t-1}, X_{1:t-1})

        Paramètres :
        - observations : tableaux des logs returns sur l'in-sample + l'OOS (pour contexte)
        - inputs : tableaux des covariables impactant les transitions sur l'in-sample + l'OOS (pour contexte)
        - oos_start : indice de début de l'OOS
        - oos_end : indice de fin de l'OOS
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")

        all_preds = self.model.filter(observations, input=inputs) # all_preds[t] = P(z_t | Y_{1:t-1}, X_{1:t-1})

        return all_preds[oos_start:oos_end]


    def get_transition_params(self):
        """
        Renvoie les paramètres des transitions (log_Ps et Ws)

        Si M = 0 (i.e. HMM classique) renvoie log_Ps, Ws = vecteur nul. Ainsi les paramètres de transitions des modèles ne sont plus que
        les intercepts (les log_Ps)
        """
        log_Ps = self.model.transitions.log_Ps.copy()
        if self.M > 0: # i.e. on a des covariables i.e. on a des poids
            Ws = self.model.transitions.Ws.copy()
        else:
            Ws = np.zeros((self.K, self.M)) # Renvoie une matrice vide (K x 0) pour ne pas casser les concaténations
        return log_Ps, Ws

    def update_transition_params(self, log_Ps : np.ndarray, Ws : np.ndarray):
        """
        Permet de modifier les paramètres de transitions (modifie les poids uniquement lorsqu'ils existent i.e. avec des covariables)

        Paramètres :
        - log_Ps : Matrice (de taille K*K) des log-probabilités de transition de base (intercepts)
        - Ws : Matrice (de taille K*M) des poids associés aux covariables macro (poids Log Regression).
        """
        self.model.transitions.log_Ps = log_Ps
        if self.M > 0: # on a des covariables donc des poids donc on met à jour
            self.model.transitions.Ws = Ws

    def compute_ll(self, Y : List[np.ndarray], X : List[np.ndarray]):
        """
        Calcule la vraisemblance sur le jeu de données en entrée

        Paramètres :
        """
        return self.model.log_likelihood(Y, inputs=X)

    def get_transition_matrices(self, X: np.ndarray = None):
        """
        Calcule les matrices de transition selon la logique demandée :
        - Si M == 0 : Modèle stationnaire, X est ignoré.
        - Si M > 0 : Modèle Macro, X est obligatoire.
        """
        if not self.is_fitted:
            raise ValueError("Modèle non fitté.")

        if self.M == 0: # HMM
            P = self.model.transitions.transition_matrix
            return P

        else: # cas M > 0 (IOHMM Macro)
            if X is None:
                raise ValueError(f"Ce modèle a été fitté avec M={self.M} variables macro, donc vous devez fournir un X pour calculer les transitions.")

            X_input = X
            T = X_input.shape[0]
            dummy_Y = np.zeros((T, self.D)) # On crée un dummy Y minimal pour satisfaire la signature technique

            return self.model.transitions.transition_matrices( # respecte la signature de ssm
                data=dummy_Y,
                input=X_input,
                mask=np.ones((T, self.D), dtype=bool),
                tag=None
            )

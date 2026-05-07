from abc import ABC, abstractmethod
import numpy as np
from typing import List, Union


class RegimeDetector(ABC):
    """
    Cette classe abstraite donne le contrat minimal que doivent respecter tous nos détecteurs de régime 

    Rôle :
    - Permettre de fitter un modèle sur des données
    - Renvoyer les probabilités d'être dans chaque régime 
    - Renvoyer les matrices de covariance dans chaque régime

    But : Eviter qu'un detecteur de régime n'implémente pas ces méthodes essentielles, ou sous d'autres noms
    
    """

    def __init__(self, n_regimes: int, n_dim: int, n_input: int = 0, random_state: int = 42):
        """
        Paramètres : 
        - n_regimes : nombre de régimes
        - n_dim : nombres d'actifs
        - n_input (optionnel) : nombre de covariables/inputs pouvant impacter les transitions
        - random_state (optionnel) : seed sur lequel on travaille

        NB : la classe en question devra rajouter l'attribut self.model
        """
        self.K = n_regimes
        self.D = n_dim
        self.M = n_input
        self.random_state = random_state
        
        self.is_fitted = False
        self.viterbi_states = None
        self.fit_observations = None
        self.fit_input = None

    @abstractmethod
    def fit(self, observations : List[np.ndarray], inputs : List[np.ndarray] = None, log_returns : List[np.ndarray] = None,
            initialize: bool = True, init_method : str = "kmeans", num_iters: int = 200, ):
        """
        Fit le modèle HMM sur les données

        Paramètres : 
        - observations : liste des tableaux des observations des HMMs (dans notre cas, une liste de taille 1)
        - inputs : list des tableaux de covariables pouvant impacter les transitions
        - log_returns : log returns de nos actifs (None si sont déjà dans observations)
        - initialize : True si on initialise les paramètres, sinon par défaut
        - init_method : méthode d'initialisation (défaut kmeans)

        """
        pass

    @abstractmethod
    def regime_probabilities(self, series_index : int = 0) -> np.ndarray:
        """
        Renvoie un tableau de probabilités (de taille (len(observations), nbr de régimes) ) d'être dans chaque régime sur le jeu d'entrainement du HMM (in-sample)

        Paramètre : 
        - series_index : index de la série sur laquelle on travaille (rappel : self.observations est une liste)
        """
        pass

    @abstractmethod
    def regime_covariances(self, log_returns: np.ndarray = None) -> Union[List[np.ndarray], np.ndarray]:
        """
        Renvoie les matrices de covariance entre les actifs par régime

        Note sur l'argument log_returns :
        - Pour les modèles statiques (ex: HMM Classique) : log_returns est ignoré. La méthode 
        renvoie les matrices de cov fixes (liste de taille K) estimées lors du fit.

        - Pour les modèles dynamiques (ex: EWMA Macro) : log_returns est obligatoire. La méthode 
        l'utilise pour calculer la séquence temporelle des matrices de cov np.ndarray (T, K, D, D).
        
        """
        pass

    @abstractmethod
    def regime_means(self, log_returns : np.ndarray = None) -> Union[List[np.ndarray], np.ndarray]:
        """
        Renvoie les moyennes des actifs par régime 

        Note sur l'argument log_returns :
        - Pour les modèles statiques (ex: HMM Classique) : log_returns est ignoré. La méthode 
        renvoie les moyennes fixes (liste de taille K) estimées lors du fit.

        - Pour les modèles dynamiques (ex: EWMA Macro) : log_returns est obligatoire. La méthode 
        l'utilise pour calculer la séquence temporelle des moyennes np.ndarray (T, K, D).
        """
        pass
    
    @abstractmethod
    def predict_probabilities(self, observations: np.ndarray, inputs: np.ndarray = None, oos_start: int = 0, oos_end: int = None) -> np.ndarray:
        """
        Calcule les probabilités prédictives sur tout l'in-sample + l'OOS (pour le contexte) mais ne renvoie les probabilités que 
        sur l'OOS

        A la ligne t, P(z_t | Y_{1:t-1}, X_{1:t-1})

        Paramètres :
        - observations : tableaux des observations du detector sur l'in-sample + l'OOS (pour contexte)
        - inputs : tableaux des covariables impactant les transitions sur l'in-sample + l'OOS (pour contexte)
        - oos_start : indice de début de l'OOS
        - oos_end : indice de fin de l'OOS
        """
        pass

    @abstractmethod
    def conditional_covariance(self, probs: np.ndarray, log_returns: np.ndarray = None) -> np.ndarray:
        """
        Renvoie les matrices de covariance en utilisant un tableau de probas 
        
        Note sur l'argument log_returns :
        - Pour les modèles statiques (ex: HMM Classique) : log_returns est ignoré. La méthode 
        utilise les paramètres fixes (liste de taille K) estimés lors du fit.

        - Pour les modèles dynamiques (ex: EWMA Macro) : log_returns est obligatoire. La méthode 
        l'utilise pour calculer la séquence temporelle des matrices np.ndarray (T, K, D, D).

        Paramètres :
        - probs : tableaux des probabilités servant à calculer les matrices de covariances à chaque date
        - log_returns : tableaux des log returns 
        """
        pass

    @abstractmethod 
    def get_params(self):
        """
        Renvoie les paramètres du modèle (Init + Trans + Obs)
        """
        pass

    @abstractmethod 
    def set_params(self, params):
        """
        Injecte des paramètres dans le modèle
        """
from abc import ABC, abstractmethod
import numpy as np
from typing import List


class RegimeDetector(ABC):
    """
    Cette classe abstraite donne le contrat minimal que doivent respecter tous nos détecteurs de régime 

    Rôle :
    - Permettre de fitter un modèle sur des données
    - Renvoyer les probabilités d'être dans chaque régime 
    - Renvoyer les matrices de covariance dans chaque régime

    But : Eviter qu'un detecteur de régime n'implémente pas ces méthodes essentielles, ou sous d'autres noms
    """

    @abstractmethod
    def fit(self, Y : List[np.ndarray], X : List[np.ndarray] = None):
        """
        Fit le modèle HMM sur les données

        Paramètres : 
        - Y : liste des tableaux de log returns (dans notre cas, une liste de taille 1)
        - X : list des tableaux de covariables pouvant impacter les transitions
        """
        pass

    @abstractmethod
    def regime_probabilities(self, series_index : int=0) -> np.ndarray:
        """
        Renvoie un tableau de probabilités (de taille (len(Y), nbr de régimes) ) d'être dans chaque régime sur le jeu d'entrainement du HMM (in-sample)

        Paramètre : 
        - series_index : index de la série sur laquelle on travaille (rappel : self.Y est une liste)
        """
        pass

    @abstractmethod
    def regime_covariances(self) -> List[np.ndarray]:
        """
        Renvoie la liste de matrices de covariance de chaque régime
        """
        pass
    
    @abstractmethod
    def predict_probabilities(self, Y: np.ndarray, X: np.ndarray = None, oos_start: int = 0, oos_end: int = None) -> np.ndarray:
        """
        Calcule les probabilités prédictives sur tout l'in-sample + l'OOS (pour le contexte) mais ne renvoie les probabilités que 
        sur l'OOS

        A la ligne t, P(z_t | Y_{1:t-1}, X_{1:t-1})

        Paramètres :
        - Y : tableaux des observations du detector sur l'in-sample + l'OOS (pour contexte)
        - X : tableaux des covariables impactant les transitions sur l'in-sample + l'OOS (pour contexte)
        - oos_start : indice de début de l'OOS
        - oos_end : indice de fin de l'OOS
        """
        pass

    @abstractmethod
    def regime_means(self) -> List[np.ndarray]:
        """
        Renvoie la liste des vecteurs de moyenne de chaque régime pour les actifs
        """
        pass

    @abstractmethod


    @abstractmethod
    def conditional_covariance(self, probs: np.ndarray) -> np.ndarray:
        """
        Renvoie les matrices de covariance en utilisant un tableau de probas

        Paramètres :
        - probs : tableaux des probabilités servant à calculer les matrices de covariances à chaque date
        """
        pass
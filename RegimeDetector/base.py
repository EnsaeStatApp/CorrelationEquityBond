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
    def predict_probabilities(self, Y: np.ndarray, X: np.ndarray = None, horizon: int = 1) -> np.ndarray:
        """
        Renvoie une série (de taille (K,) ) de probabilités d'être dans chaque régime à l'instant T + horizon sachant tout jusqu'à len(Y) 
        i.e. le vecteur P(z_{T+horizon} | Y_{1:T}, X_{1:T})
        
        Paramètres :
        - Y : tableau de log returns
        - X : tableau de covariables pouvant impacter les transitions
        - horizon : tq que T+horizon est la date à laquelle on calcule les probabilités

        NB : cette fonction a besoin de prendre le Y et le X en entier pour le contexte de la probabilité 
        """
        pass


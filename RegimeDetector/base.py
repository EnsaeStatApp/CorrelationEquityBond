from abc import ABC, abstractmethod
import numpy as np


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
    def fit(self, Y : list, X : list=None):
        pass

    @abstractmethod
    def regime_probabilities(self, series_index : int=0, Y: np.ndarray=None, X: np.ndarray=None):
        """
        Si Y = None (et a priori X = None) alors calcule les probabilités sur les données de fit 
        Sinon sur l'out-of-sample
        
        """
        pass

    @abstractmethod
    def regime_covariances(self):
        """
        Σ_k : covariance par régime
        """
        pass

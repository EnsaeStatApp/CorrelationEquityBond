from abc import ABC, abstractmethod
import numpy as np
from RegimeDetector.base import RegimeDetector
from Allocator.base import Allocator

class Strategy(ABC):
    """
    Classe abstraite définissant la structure que doivent respecter nos stratégies d'allocation dynamique 

    Une Strategy :
    - utilise un RegimeDetector
    - utilise un Allocator
    - produit une série temporelle de poids
    - calcule la performance
    """

    def __init__(self, detector : RegimeDetector, allocator : Allocator):
        """
        detector : Modèle de détection de régimes 
        allocator : Méthode d'allocation basée sur une matrice de covariance 
            
        """
        self.detector = detector
        self.allocator = allocator

        self.weights_ = None
        self.returns_ = None
        self.portfolio_returns_ = None

    @abstractmethod
    def compute_covariance(self, t, Y, X=None):
        """
        Détermine la matrice de covariance utilisée à la date t (en utilisant seulement les datas jusqu'en t-1)

        Plusieurs possibilités 
        - Cov globale (benchmark pour faire VS tous nos markov models)
        - Cov régime (Sigma_k et dans ce cas prendre le regime k tq k = argmax [k-> p_k(t)])       
        - Cov pondérée par probabilités (Sigma_t = sum_k [p_k(t) Sigma_k])
        - ... 
        """
        pass

    def run(self, Y, X=None):
        """
        Exécute la stratégie à chaque date

        Y : Rendements des actifs
        X : Covariables macro (pour IOHMM)
            
        """
        T, D = Y.shape
        weights = np.zeros((T, D))
        portfolio_returns = np.zeros(T)

        for t in range(T):
            Sigma_t = self.compute_covariance(t, Y, X) # recup la matrice de cov au temps t choisie
            w_t = self.allocator.allocate(Sigma_t) # recup les poids donnés par la méthode d'allocation choisie

            weights[t] = w_t # stocke ces poids dans un vecteur 
            portfolio_returns[t] = w_t @ Y[t] # stocke le return du portefeuille en t avec les poids w_t 

        self.weights_ = weights
        self.returns_ = Y
        self.portfolio_returns_ = portfolio_returns

        return portfolio_returns
    
    def performance_summary(self):
        """
        Calcule quelques métriques standards. 
        TO DO: rajouter celles qu'on veut pour notre sujet
        """
        r = self.portfolio_returns_

        sharpe = np.mean(r) / (np.std(r) + 1e-8) * np.sqrt(12)
        vol = np.std(r) * np.sqrt(12)
        cum = np.cumprod(1 + r)
        mdd = np.min(cum / np.maximum.accumulate(cum) - 1)

        return {"Sharpe": sharpe, "Volatility": vol, "MaxDrawdown": mdd}



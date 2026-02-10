from abc import ABC, abstractmethod
import numpy as np


class Allocator(ABC):
    """
    Classe abstraite qui donne le contrat minimal que doivent respecter tous nos allocateurs d'actifs
    i.e. renvoyer le vecteur des poids en suivant une méthode d'allocation à partir d'une matrice de cov
    """
    @abstractmethod
    def allocate(self, Sigma : np.ndarray) -> np.ndarray:
        """
        Sigma : matrice de cov de taille D*D où D est le nbr d'actifs
        Renvoie le vecteur des poids (de taille D) en suivant une méthode particulière
    
        """
        pass

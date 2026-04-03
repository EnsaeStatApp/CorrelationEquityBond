import numpy as np
from RegimeDetector.base import RegimeDetector


class BaseRegimeDetector(RegimeDetector):
    """
    Cette classe (abstraite toujours) donne le comportement commun de chacun de nos détecteurs de régimes. 
    
    But : Ne pas réimplémenter à chaque fois les mêmes fonctions. 
    """
    def regime_correlations(self):
        """
        Corrélations par régime : 
        regime_correlations[k] = matrice de corrélations du régime k 
        """
        corrs = []
        for S in self.regime_covariances():
            d = np.sqrt(np.diag(S)) # vecteurs des vols pour ce régime
            product_vols_matrix = np.outer(d, d) # product_vols_matrix[i,j] = sigma[i] sigma[j]
            corrs.append(S / product_vols_matrix) 
        return corrs
